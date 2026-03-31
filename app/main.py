from flask import Flask, jsonify, request, send_file, session, send_from_directory
from io import BytesIO
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os, json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "db.json")

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path='')
app.secret_key = "segredo"

APP_PIN = "1234"

# =========================
# DB
# =========================
def ensure_data():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(DB_PATH):
        with open(DB_PATH, "w", encoding="utf-8") as f:
            json.dump({
                "purchases": [],
                "cash_state": {},
                "next_id": 1
            }, f)

def load_db():
    ensure_data()
    with open(DB_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_db(db):
    with open(DB_PATH, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2)

# =========================
# FRONTEND
# =========================
@app.route("/")
def home():
    return send_from_directory(STATIC_DIR, "index.html")

@app.route("/<path:path>")
def static_files(path):
    full = os.path.join(STATIC_DIR, path)
    if os.path.exists(full):
        return send_from_directory(STATIC_DIR, path)
    return send_from_directory(STATIC_DIR, "index.html")

# =========================
# AUTH
# =========================
def auth():
    if not session.get("ok"):
        return jsonify({"error": "login"}), 401

@app.route("/api/login", methods=["POST"])
def login():
    if request.json.get("pin") == APP_PIN:
        session["ok"] = True
        return jsonify({"ok": True})
    return jsonify({"ok": False})

@app.route("/api/status")
def status():
    return jsonify({"logged_in": bool(session.get("ok"))})

@app.route("/api/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"ok": True})

# =========================
# COMPRAS
# =========================
@app.route("/api/purchases", methods=["GET", "POST"])
def purchases():
    if auth(): return auth()
    db = load_db()

    if request.method == "GET":
        return jsonify(db["purchases"])

    item = request.json
    item["id"] = db["next_id"]
    db["next_id"] += 1
    db["purchases"].append(item)
    save_db(db)

    return jsonify({"ok": True})

@app.route("/api/purchases/<int:id>/complete", methods=["POST"])
def complete(id):
    db = load_db()
    for i in db["purchases"]:
        if i["id"] == id:
            i["qty_bought"] = i.get("qty_to_buy", 0)
    save_db(db)
    return jsonify({"ok": True})

@app.route("/api/purchases/<int:id>", methods=["DELETE"])
def delete(id):
    db = load_db()
    db["purchases"] = [i for i in db["purchases"] if i["id"] != id]
    save_db(db)
    return jsonify({"ok": True})

# =========================
# CAIXA
# =========================
@app.route("/api/cash-state", methods=["GET", "POST"])
def cash():
    if auth(): return auth()
    db = load_db()

    if request.method == "GET":
        return jsonify(db.get("cash_state", {}))

    db["cash_state"] = request.json
    save_db(db)
    return jsonify({"ok": True})

# =========================
# EXPORT EXCEL
# =========================
@app.route("/api/export/purchases/excel")
def exp_excel():
    db = load_db()
    wb = Workbook()
    ws = wb.active

    ws.append(["Artigo", "Qtd"])

    for p in db["purchases"]:
        ws.append([p.get("name"), p.get("qty_to_buy")])

    mem = BytesIO()
    wb.save(mem)
    mem.seek(0)

    return send_file(mem, as_attachment=True, download_name="compras.xlsx")

# =========================
# EXPORT PDF
# =========================
@app.route("/api/export/purchases/pdf")
def exp_pdf():
    db = load_db()

    mem = BytesIO()
    c = canvas.Canvas(mem, pagesize=A4)

    y = 800
    for p in db["purchases"]:
        c.drawString(50, y, f"{p.get('name')} - {p.get('qty_to_buy')}")
        y -= 20

    c.save()
    mem.seek(0)

    return send_file(mem, as_attachment=True, download_name="compras.pdf")

# =========================
# START
# =========================
if __name__ == "__main__":
    ensure_data()
    app.run(host="0.0.0.0", port=5000)