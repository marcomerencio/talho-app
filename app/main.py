from flask import Flask, send_from_directory, request, jsonify, session, send_file
import os
import json
import threading
import unicodedata
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook, Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, 'static')
DATA_DIR = os.path.join(BASE_DIR, 'data')
EXPORTS_DIR = os.path.join(BASE_DIR, 'exports')
PURCHASES_EXPORT_DIR = os.path.join(EXPORTS_DIR, 'compras')
CASH_EXPORT_DIR = os.path.join(EXPORTS_DIR, 'fechos_caixa')

DB_PATH = os.path.join(DATA_DIR, 'db.json')
EXCEL_PATH = os.path.join(DATA_DIR, 'base_sage.xlsx')
LOCK = threading.Lock()

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path='')
app.secret_key = os.environ.get('SECRET_KEY', 'talho-secret-key-change-me')
APP_PIN = os.environ.get('APP_PIN', '1234')

DEFAULT_DB = {
    'stock': [],
    'purchases': [],
    'clients': [],
    'cash_state': {
        'talho': {
            'date': '',
            'start': 100,
            'inCash': 0,
            'inMb': 0,
            'inMbway': 0,
            'inOther': 0,
            'out': 0,
            'obs': '',
            'notes': {'500': 0, '200': 0, '100': 0, '50': 0, '20': 0, '10': 0, '5': 0},
            'coins': {'2': 0, '1': 0, '0.5': 0, '0.2': 0, '0.1': 0, '0.05': 0, '0.02': 0, '0.01': 0}
        },
        'cong': {
            'date': '',
            'start': 100,
            'inCash': 0,
            'inMb': 0,
            'inMbway': 0,
            'inOther': 0,
            'out': 0,
            'obs': '',
            'notes': {'500': 0, '200': 0, '100': 0, '50': 0, '20': 0, '10': 0, '5': 0},
            'coins': {'2': 0, '1': 0, '0.5': 0, '0.2': 0, '0.1': 0, '0.05': 0, '0.02': 0, '0.01': 0}
        }
    },
    'next_ids': {
        'purchase': 1
    }
}


def ensure_dirs() -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(PURCHASES_EXPORT_DIR, exist_ok=True)
    os.makedirs(CASH_EXPORT_DIR, exist_ok=True)


def ensure_db() -> None:
    ensure_dirs()
    if not os.path.exists(DB_PATH):
        with open(DB_PATH, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_DB, f, ensure_ascii=False, indent=2)


def load_db() -> dict:
    ensure_db()
    with LOCK:
        with open(DB_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)

    changed = False

    for key, value in DEFAULT_DB.items():
        if key not in data:
            data[key] = value
            changed = True

    if 'next_ids' not in data:
        data['next_ids'] = DEFAULT_DB['next_ids']
        changed = True
    else:
        if 'purchase' not in data['next_ids']:
            data['next_ids']['purchase'] = 1
            changed = True

    if changed:
        save_db(data)

    return data


def save_db(data: dict) -> None:
    ensure_dirs()
    with LOCK:
        with open(DB_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


def require_login():
    if not session.get('logged_in'):
        return jsonify({'error': 'Não autenticado'}), 401
    return None


def parse_amount(value) -> float:
    if value is None or value == '':
        return 0.0
    return round(float(str(value).replace(',', '.')), 2)


def normalize_text(value):
    text = str(value or '').strip().lower()
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
    return text


def find_sheet_name(workbook, expected_name):
    expected = normalize_text(expected_name)
    for sheet_name in workbook.sheetnames:
        if normalize_text(sheet_name) == expected:
            return sheet_name
    return None


def load_excel_master():
    if not os.path.exists(EXCEL_PATH):
        return {'articles': [], 'suppliers': []}

    wb = load_workbook(EXCEL_PATH, data_only=True)
    articles = []
    suppliers = []

    artigos_sheet = find_sheet_name(wb, 'artigos')
    fornecedores_sheet = find_sheet_name(wb, 'fornecedores')

    if artigos_sheet:
        ws = wb[artigos_sheet]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            headers = [normalize_text(h) for h in rows[0]]
            for row in rows[1:]:
                item = dict(zip(headers, row))
                code = str(item.get('codigo', '') or '').strip()
                name = str(item.get('nome', '') or '').strip()
                if code and name:
                    articles.append({'code': code, 'name': name})

    if fornecedores_sheet:
        ws = wb[fornecedores_sheet]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            headers = [normalize_text(h) for h in rows[0]]
            for row in rows[1:]:
                item = dict(zip(headers, row))
                code = str(item.get('codigo', '') or '').strip()
                name = str(item.get('nome', '') or '').strip()
                if code and name:
                    suppliers.append({'code': code, 'name': name})

    return {'articles': articles, 'suppliers': suppliers}


def purchase_state(item: dict) -> str:
    qty_to_buy = parse_amount(item.get('qty_to_buy', 0))
    qty_bought = parse_amount(item.get('qty_bought', 0))
    if qty_bought <= 0:
        return 'Por comprar'
    if qty_bought < qty_to_buy:
        return 'Parcial'
    return 'Comprado'


def calc_cash_summary(section_data: dict) -> dict:
    notes_total = 0.0
    coins_total = 0.0

    note_values = {'500': 500, '200': 200, '100': 100, '50': 50, '20': 20, '10': 10, '5': 5}
    coin_values = {'2': 2, '1': 1, '0.5': 0.5, '0.2': 0.2, '0.1': 0.1, '0.05': 0.05, '0.02': 0.02, '0.01': 0.01}

    for key, value in note_values.items():
        notes_total += parse_amount(section_data.get('notes', {}).get(key, 0)) * value

    for key, value in coin_values.items():
        coins_total += parse_amount(section_data.get('coins', {}).get(key, 0)) * value

    real = round(notes_total + coins_total, 2)
    expected = round(
        parse_amount(section_data.get('start', 0))
        + parse_amount(section_data.get('inCash', 0))
        + parse_amount(section_data.get('inMb', 0))
        + parse_amount(section_data.get('inMbway', 0))
        + parse_amount(section_data.get('inOther', 0))
        - parse_amount(section_data.get('out', 0)),
        2
    )
    diff = round(real - expected, 2)

    status = 'Certo'
    if abs(diff) >= 0.01:
        status = 'Sobra em caixa' if diff > 0 else 'Falta em caixa'

    return {
        'notes_total': round(notes_total, 2),
        'coins_total': round(coins_total, 2),
        'real': real,
        'expected': expected,
        'diff': diff,
        'status': status
    }


def save_workbook_to_disk_and_memory(workbook: Workbook, filepath: str):
    ensure_dirs()
    workbook.save(filepath)
    mem = BytesIO()
    workbook.save(mem)
    mem.seek(0)
    return mem


def wrap_pdf_text(c: canvas.Canvas, text: str, x: int, y: int, max_chars: int = 95, line_height: int = 16):
    text = str(text or '')
    while text:
        chunk = text[:max_chars]
        c.drawString(x, y, chunk)
        y -= line_height
        text = text[max_chars:]
    return y


@app.route('/')
def index():
    return send_from_directory(STATIC_DIR, 'index.html')


@app.route('/api/login', methods=['POST'])
def api_login():
    data = request.get_json(silent=True) or {}
    pin = str(data.get('pin', '')).strip()

    if pin == APP_PIN:
        session['logged_in'] = True
        return jsonify({'ok': True}), 200

    return jsonify({'ok': False, 'error': 'PIN inválido'}), 401


@app.route('/api/logout', methods=['POST'])
def api_logout():
    session.clear()
    return jsonify({'ok': True}), 200


@app.route('/api/status')
def api_status():
    return jsonify({'logged_in': bool(session.get('logged_in'))})


@app.route('/api/master/articles')
def api_master_articles():
    auth = require_login()
    if auth:
        return auth
    return jsonify(load_excel_master()['articles'])


@app.route('/api/master/suppliers')
def api_master_suppliers():
    auth = require_login()
    if auth:
        return auth
    return jsonify(load_excel_master()['suppliers'])


@app.route('/api/stock', methods=['GET', 'POST'])
def api_stock():
    auth = require_login()
    if auth:
        return auth

    data = load_db()

    if request.method == 'GET':
        return jsonify(data.get('stock', []))

    payload = request.get_json(silent=True) or {}
    item = {
        'name': str(payload.get('name', '')).strip(),
        'qty': parse_amount(payload.get('qty', 0)),
        'unit': str(payload.get('unit', 'kg')).strip(),
        'min': parse_amount(payload.get('min', 0))
    }

    if not item['name']:
        return jsonify({'error': 'Produto obrigatório'}), 400

    data['stock'].append(item)
    save_db(data)
    return jsonify(item), 201


@app.route('/api/purchases', methods=['GET', 'POST'])
def api_purchases():
    auth = require_login()
    if auth:
        return auth

    data = load_db()

    if request.method == 'GET':
        return jsonify(data.get('purchases', []))

    payload = request.get_json(silent=True) or {}
    item = {
        'id': data['next_ids']['purchase'],
        'code': str(payload.get('code', '')).strip(),
        'name': str(payload.get('name', '')).strip(),
        'supplier_code': str(payload.get('supplier_code', '')).strip(),
        'supplier': str(payload.get('supplier', '')).strip(),
        'qty_to_buy': parse_amount(payload.get('qty_to_buy', 0)),
        'qty_bought': parse_amount(payload.get('qty_bought', 0)),
        'unit': str(payload.get('unit', 'kg')).strip(),
        'priority': str(payload.get('priority', 'Média')).strip()
    }

    if not item['code'] or not item['name']:
        return jsonify({'error': 'Artigo obrigatório'}), 400

    data['next_ids']['purchase'] += 1
    data['purchases'].append(item)
    save_db(data)
    return jsonify(item), 201


@app.route('/api/purchases/<int:item_id>/complete', methods=['POST'])
def api_purchase_complete(item_id: int):
    auth = require_login()
    if auth:
        return auth

    data = load_db()

    for item in data.get('purchases', []):
        if item.get('id') == item_id:
            item['qty_bought'] = item.get('qty_to_buy', 0)
            save_db(data)
            return jsonify({'ok': True, 'item': item}), 200

    return jsonify({'error': 'Artigo não encontrado'}), 404


@app.route('/api/purchases/<int:item_id>', methods=['DELETE'])
def api_purchase_delete(item_id: int):
    auth = require_login()
    if auth:
        return auth

    data = load_db()
    before = len(data.get('purchases', []))
    data['purchases'] = [p for p in data.get('purchases', []) if p.get('id') != item_id]

    if len(data['purchases']) == before:
        return jsonify({'error': 'Artigo não encontrado'}), 404

    save_db(data)
    return jsonify({'ok': True}), 200


@app.route('/api/clients', methods=['GET', 'POST'])
def api_clients():
    auth = require_login()
    if auth:
        return auth

    data = load_db()

    if request.method == 'GET':
        return jsonify(data.get('clients', []))

    payload = request.get_json(silent=True) or {}
    item = {
        'name': str(payload.get('name', '')).strip(),
        'phone': str(payload.get('phone', '')).strip(),
        'note': str(payload.get('note', '')).strip()
    }

    if not item['name']:
        return jsonify({'error': 'Nome obrigatório'}), 400

    data['clients'].append(item)
    save_db(data)
    return jsonify(item), 201


@app.route('/api/cash-state', methods=['GET', 'POST'])
def api_cash_state():
    auth = require_login()
    if auth:
        return auth

    data = load_db()

    if request.method == 'GET':
        return jsonify(data.get('cash_state', DEFAULT_DB['cash_state']))

    payload = request.get_json(silent=True) or {}
    talho = payload.get('talho', DEFAULT_DB['cash_state']['talho'])
    cong = payload.get('cong', DEFAULT_DB['cash_state']['cong'])

    data['cash_state'] = {
        'talho': talho,
        'cong': cong
    }
    save_db(data)
    return jsonify({'ok': True}), 200


@app.route('/api/export/purchases/excel')
def api_export_purchases_excel():
    auth = require_login()
    if auth:
        return auth

    data = load_db()
    wb = Workbook()
    ws = wb.active
    ws.title = 'compras'

    ws.append([
        'id', 'codigo_artigo', 'artigo', 'codigo_fornecedor', 'fornecedor',
        'quantidade_a_comprar', 'quantidade_comprada', 'unidade', 'prioridade', 'estado'
    ])

    for item in data.get('purchases', []):
        ws.append([
            item.get('id'),
            item.get('code'),
            item.get('name'),
            item.get('supplier_code'),
            item.get('supplier'),
            item.get('qty_to_buy'),
            item.get('qty_bought'),
            item.get('unit'),
            item.get('priority'),
            purchase_state(item)
        ])

    filename = f"compras-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.xlsx"
    filepath = os.path.join(PURCHASES_EXPORT_DIR, filename)
    mem = save_workbook_to_disk_and_memory(wb, filepath)

    return send_file(
        mem,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/export/purchases/pdf')
def api_export_purchases_pdf():
    auth = require_login()
    if auth:
        return auth

    data = load_db()
    filename = f"compras-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.pdf"
    filepath = os.path.join(PURCHASES_EXPORT_DIR, filename)

    mem = BytesIO()
    c = canvas.Canvas(mem, pagesize=A4)
    width, height = A4
    y = height - 40

    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "GRUPO BOLHÃO - Artigos a Comprar")
    y -= 25

    c.setFont("Helvetica", 10)
    c.drawString(40, y, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    y -= 25

    for item in data.get('purchases', []):
        if y < 100:
            c.showPage()
            y = height - 40
            c.setFont("Helvetica", 10)

        estado = purchase_state(item)
        lines = [
            f"Código: {item.get('code', '')} | Artigo: {item.get('name', '')}",
            f"Fornecedor: {item.get('supplier_code', '')} - {item.get('supplier', '')}",
            f"Qtd comprar: {item.get('qty_to_buy', 0)} | Qtd comprada: {item.get('qty_bought', 0)} | Unidade: {item.get('unit', '')}",
            f"Prioridade: {item.get('priority', '')} | Estado: {estado}"
        ]

        c.setFont("Helvetica-Bold", 10)
        c.drawString(40, y, f"Artigo #{item.get('id', '')}")
        y -= 16
        c.setFont("Helvetica", 10)

        for line in lines:
            y = wrap_pdf_text(c, line, 50, y, 95, 14)

        y -= 8
        c.line(40, y, width - 40, y)
        y -= 14

    c.save()
    mem.seek(0)

    with open(filepath, 'wb') as f:
        f.write(mem.getvalue())

    mem.seek(0)
    return send_file(mem, as_attachment=True, download_name=filename, mimetype='application/pdf')


@app.route('/api/export/cash/excel')
def api_export_cash_excel():
    auth = require_login()
    if auth:
        return auth

    data = load_db()
    cash_state = data.get('cash_state', DEFAULT_DB['cash_state'])

    wb = Workbook()
    ws = wb.active
    ws.title = 'fecho_caixa'

    ws.append([
        'secao', 'data', 'fundo_inicial', 'entradas_dinheiro', 'entradas_multibanco',
        'entradas_mbway', 'outras_entradas', 'saidas',
        'total_notas', 'total_moedas', 'caixa_real', 'saldo_teorico', 'diferenca', 'estado', 'observacoes'
    ])

    for section_name, label in [('talho', 'Talho'), ('cong', 'Congelados')]:
        section_data = cash_state.get(section_name, {})
        summary = calc_cash_summary(section_data)
        ws.append([
            label,
            section_data.get('date', ''),
            section_data.get('start', 0),
            section_data.get('inCash', 0),
            section_data.get('inMb', 0),
            section_data.get('inMbway', 0),
            section_data.get('inOther', 0),
            section_data.get('out', 0),
            summary['notes_total'],
            summary['coins_total'],
            summary['real'],
            summary['expected'],
            summary['diff'],
            summary['status'],
            section_data.get('obs', '')
        ])

    filename = f"fecho-caixa-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.xlsx"
    filepath = os.path.join(CASH_EXPORT_DIR, filename)
    mem = save_workbook_to_disk_and_memory(wb, filepath)

    return send_file(
        mem,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/export/cash/pdf')
def api_export_cash_pdf():
    auth = require_login()
    if auth:
        return auth

    data = load_db()
    cash_state = data.get('cash_state', DEFAULT_DB['cash_state'])

    filename = f"fecho-caixa-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.pdf"
    filepath = os.path.join(CASH_EXPORT_DIR, filename)

    mem = BytesIO()
    c = canvas.Canvas(mem, pagesize=A4)
    width, height = A4
    y = height - 40

    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "GRUPO BOLHÃO - Fecho de Caixa")
    y -= 25

    c.setFont("Helvetica", 10)
    c.drawString(40, y, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    y -= 30

    for section_name, label in [('talho', 'Talho'), ('cong', 'Congelados')]:
        section_data = cash_state.get(section_name, {})
        summary = calc_cash_summary(section_data)

        if y < 180:
            c.showPage()
            y = height - 40

        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, y, f"Secção: {label}")
        y -= 18

        c.setFont("Helvetica", 10)
        lines = [
            f"Data: {section_data.get('date', '')}",
            f"Fundo inicial: {section_data.get('start', 0)} €",
            f"Entradas dinheiro: {section_data.get('inCash', 0)} €",
            f"Entradas multibanco: {section_data.get('inMb', 0)} €",
            f"Entradas MBWay: {section_data.get('inMbway', 0)} €",
            f"Outras entradas: {section_data.get('inOther', 0)} €",
            f"Saídas: {section_data.get('out', 0)} €",
            f"Total notas: {summary['notes_total']} €",
            f"Total moedas: {summary['coins_total']} €",
            f"Caixa real: {summary['real']} €",
            f"Saldo teórico: {summary['expected']} €",
            f"Diferença: {summary['diff']} €",
            f"Estado: {summary['status']}",
            f"Observações: {section_data.get('obs', '')}"
        ]

        for line in lines:
            y = wrap_pdf_text(c, line, 50, y, 95, 14)

        y -= 8
        c.line(40, y, width - 40, y)
        y -= 18

    c.save()
    mem.seek(0)

    with open(filepath, 'wb') as f:
        f.write(mem.getvalue())

    mem.seek(0)
    return send_file(mem, as_attachment=True, download_name=filename, mimetype='application/pdf')


@app.route('/<path:path>')
def serve_static(path: str):
    file_path = os.path.join(STATIC_DIR, path)
    if os.path.exists(file_path):
        return send_from_directory(STATIC_DIR, path)
    return send_from_directory(STATIC_DIR, 'index.html')


if __name__ == '__main__':
    ensure_db()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))