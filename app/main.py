from __future__ import annotations
from flask import Flask, send_from_directory, request, jsonify, session
import os
import json
from datetime import datetime, date
from io import BytesIO
from threading import Lock
import zipfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, 'static')
DATA_DIR = os.path.join(BASE_DIR, 'data')
DB_PATH = os.path.join(DATA_DIR, 'db.json')
LOCK = Lock()

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path='')
app.secret_key = os.environ.get('SECRET_KEY', 'talho-secret-key-change-me')
APP_PIN = os.environ.get('APP_PIN', '1234')

DEFAULT_DB = {
    'transactions': [],
    'closures': [],
    'suppliers': [
        {'id': 1, 'name': 'Fornecedor Exemplo', 'contact': '', 'notes': 'Editar ou apagar.'}
    ],
    'next_ids': {
        'transaction': 1,
        'closure': 1,
        'supplier': 2
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


def today_iso() -> str:
    return date.today().isoformat()


def parse_amount(value) -> float:
    if value is None or value == '':
        return 0.0
    return round(float(str(value).replace(',', '.')), 2)


def get_today_transactions(data: dict, day: str | None = None) -> list[dict]:
    day = day or today_iso()
    return [t for t in data['transactions'] if t.get('date') == day]


def compute_summary(transactions: list[dict]) -> dict:
    sales_talho = 0.0
    sales_congelados = 0.0
    expenses = 0.0
    other_entries = 0.0
    by_payment = {'dinheiro': 0.0, 'multibanco': 0.0, 'mbway': 0.0, 'outro': 0.0}
    for t in transactions:
        amount = parse_amount(t.get('amount', 0))
        ttype = t.get('type', 'sale')
        section = t.get('section', 'talho')
        payment = (t.get('payment_method') or 'dinheiro').lower()
        if ttype == 'sale':
            if section == 'congelados':
                sales_congelados += amount
            else:
                sales_talho += amount
            if payment not in by_payment:
                payment = 'outro'
            by_payment[payment] += amount
        elif ttype == 'expense':
            expenses += amount
        elif ttype == 'entry':
            other_entries += amount

    total_sales = sales_talho + sales_congelados
    gross_cash_expected = by_payment['dinheiro'] + other_entries - expenses
    return {
        'sales_talho': round(sales_talho, 2),
        'sales_congelados': round(sales_congelados, 2),
        'total_sales': round(total_sales, 2),
        'expenses': round(expenses, 2),
        'other_entries': round(other_entries, 2),
        'by_payment': {k: round(v, 2) for k, v in by_payment.items()},
        'cash_expected': round(gross_cash_expected, 2),
        'count': len(transactions)
    }


def note_counts_total(note_counts: dict) -> float:
    values = {
        '500': 500.0, '200': 200.0, '100': 100.0, '50': 50.0, '20': 20.0,
        '10': 10.0, '5': 5.0, '2': 2.0, '1': 1.0,
        '0.50': 0.5, '0.20': 0.2, '0.10': 0.1, '0.05': 0.05,
        '0.02': 0.02, '0.01': 0.01,
    }
    total = 0.0
    note_counts = note_counts or {}
    for k, v in values.items():
        qty = int(note_counts.get(k, 0) or 0)
        total += qty * v
    return round(total, 2)


@app.route('/')
def index():
    return send_from_directory(STATIC_DIR, 'index.html')


@app.route('/health')
def health():
    return {'status': 'ok'}, 200


@app.route('/api/login', methods=['POST'])
def login():
    payload = request.get_json(silent=True) or {}
    pin = str(payload.get('pin', ''))
    if pin == APP_PIN:
        session['logged_in'] = True
        return jsonify({'ok': True})
    return jsonify({'ok': False, 'error': 'PIN inválido'}), 401


@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'ok': True})


@app.route('/api/status')
def status():
    return jsonify({'logged_in': bool(session.get('logged_in'))})


@app.route('/api/summary')
def summary():
    auth = require_login()
    if auth:
        return auth
    day = request.args.get('date') or today_iso()
    data = load_db()
    transactions = get_today_transactions(data, day)
    result = compute_summary(transactions)
    result['date'] = day
    return jsonify(result)


@app.route('/api/transactions', methods=['GET', 'POST'])
def transactions():
    auth = require_login()
    if auth:
        return auth
    data = load_db()
    if request.method == 'GET':
        day = request.args.get('date') or today_iso()
        tx = sorted(get_today_transactions(data, day), key=lambda x: x.get('created_at', ''), reverse=True)
        return jsonify(tx)

    payload = request.get_json(silent=True) or {}
    ttype = payload.get('type', 'sale')
    amount = parse_amount(payload.get('amount', 0))
    if amount <= 0:
        return jsonify({'error': 'Valor inválido'}), 400

    tx = {
        'id': data['next_ids']['transaction'],
        'date': payload.get('date') or today_iso(),
        'type': ttype,
        'section': payload.get('section', 'talho'),
        'amount': amount,
        'payment_method': payload.get('payment_method', 'dinheiro'),
        'description': (payload.get('description') or '').strip(),
        'created_at': datetime.now().isoformat(timespec='seconds')
    }
    data['next_ids']['transaction'] += 1
    data['transactions'].append(tx)
    save_db(data)
    return jsonify(tx), 201


@app.route('/api/transactions/<int:tx_id>', methods=['DELETE'])
def delete_transaction(tx_id: int):
    auth = require_login()
    if auth:
        return auth
    data = load_db()
    before = len(data['transactions'])
    data['transactions'] = [t for t in data['transactions'] if t.get('id') != tx_id]
    if len(data['transactions']) == before:
        return jsonify({'error': 'Movimento não encontrado'}), 404
    save_db(data)
    return jsonify({'ok': True})


@app.route('/api/suppliers', methods=['GET', 'POST'])
def suppliers():
    auth = require_login()
    if auth:
        return auth
    data = load_db()
    if request.method == 'GET':
        return jsonify(data['suppliers'])

    payload = request.get_json(silent=True) or {}
    name = (payload.get('name') or '').strip()
    if not name:
        return jsonify({'error': 'Nome obrigatório'}), 400
    supplier = {
        'id': data['next_ids']['supplier'],
        'name': name,
        'contact': (payload.get('contact') or '').strip(),
        'notes': (payload.get('notes') or '').strip(),
    }
    data['next_ids']['supplier'] += 1
    data['suppliers'].append(supplier)
    save_db(data)
    return jsonify(supplier), 201


@app.route('/api/suppliers/<int:supplier_id>', methods=['DELETE'])
def delete_supplier(supplier_id: int):
    auth = require_login()
    if auth:
        return auth
    data = load_db()
    before = len(data['suppliers'])
    data['suppliers'] = [s for s in data['suppliers'] if s.get('id') != supplier_id]
    if len(data['suppliers']) == before:
        return jsonify({'error': 'Fornecedor não encontrado'}), 404
    save_db(data)
    return jsonify({'ok': True})


@app.route('/api/close-day', methods=['POST'])
def close_day():
    auth = require_login()
    if auth:
        return auth
    data = load_db()
    payload = request.get_json(silent=True) or {}
    day = payload.get('date') or today_iso()
    note_counts = payload.get('note_counts') or {}
    observed_cash = note_counts_total(note_counts)
    summary = compute_summary(get_today_transactions(data, day))
    difference = round(observed_cash - summary['cash_expected'], 2)

    closure = {
        'id': data['next_ids']['closure'],
        'date': day,
        'summary': summary,
        'note_counts': note_counts,
        'observed_cash': observed_cash,
        'difference': difference,
        'notes': (payload.get('notes') or '').strip(),
        'created_at': datetime.now().isoformat(timespec='seconds')
    }
    data['next_ids']['closure'] += 1
    data['closures'].append(closure)
    save_db(data)
    return jsonify(closure), 201


@app.route('/api/history')
def history():
    auth = require_login()
    if auth:
        return auth
    data = load_db()
    closures = sorted(data['closures'], key=lambda x: x.get('created_at', ''), reverse=True)
    return jsonify(closures)


@app.route('/api/backup')
def backup():
    auth = require_login()
    if auth:
        return auth
    ensure_db()
    mem = BytesIO()
    with zipfile.ZipFile(mem, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(DB_PATH, arcname='backup/db.json')
    mem.seek(0)
    filename = f"backup-app-talho-{today_iso()}.zip"
    return send_file(mem, as_attachment=True, download_name=filename, mimetype='application/zip')


@app.route('/<path:path>')
def serve_static(path: str):
    file_path = os.path.join(STATIC_DIR, path)
    if os.path.exists(file_path):
        return send_from_directory(STATIC_DIR, path)
    return send_from_directory(STATIC_DIR, 'index.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
