from flask import Flask, send_from_directory, request, jsonify, session
import os
import json
import threading
from datetime import date
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, 'static')
DATA_DIR = os.path.join(BASE_DIR, 'data')
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
    }
}


def ensure_db() -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
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

    if changed:
        save_db(data)

    return data


def save_db(data: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
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


def load_excel_master():
    if not os.path.exists(EXCEL_PATH):
        return {'articles': [], 'suppliers': []}

    wb = load_workbook(EXCEL_PATH, data_only=True)
    articles = []
    suppliers = []

    if 'artigos' in wb.sheetnames:
        ws = wb['artigos']
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            headers = [str(h).strip().lower() if h is not None else '' for h in rows[0]]
            for row in rows[1:]:
                item = dict(zip(headers, row))
                code = str(item.get('codigo', '') or '').strip()
                name = str(item.get('nome', '') or '').strip()
                if code or name:
                    articles.append({
                        'code': code,
                        'name': name
                    })

    if 'fornecedores' in wb.sheetnames:
        ws = wb['fornecedores']
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            headers = [str(h).strip().lower() if h is not None else '' for h in rows[0]]
            for row in rows[1:]:
                item = dict(zip(headers, row))
                code = str(item.get('codigo', '') or '').strip()
                name = str(item.get('nome', '') or '').strip()
                if code or name:
                    suppliers.append({
                        'code': code,
                        'name': name
                    })

    return {'articles': articles, 'suppliers': suppliers}


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

    data['purchases'].append(item)
    save_db(data)
    return jsonify(item), 201


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


@app.route('/<path:path>')
def serve_static(path: str):
    file_path = os.path.join(STATIC_DIR, path)
    if os.path.exists(file_path):
        return send_from_directory(STATIC_DIR, path)
    return send_from_directory(STATIC_DIR, 'index.html')


if __name__ == '__main__':
    ensure_db()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))