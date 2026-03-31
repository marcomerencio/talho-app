from flask import Flask, jsonify, request, send_file
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)

# 🔥 BASE DE DADOS EM MEMÓRIA (simples)
purchases = []
cash_state = {
    "talho": {"start": 0, "in": 0, "out": 0},
    "cong": {"start": 0, "in": 0, "out": 0}
}

# =========================
# 🛒 COMPRAS
# =========================

@app.route("/api/purchases", methods=["GET", "POST"])
def api_purchases():
    global purchases

    if request.method == "GET":
        return jsonify(purchases)

    data = request.json
    purchases.append(data)
    return jsonify({"ok": True})


# =========================
# 💰 CAIXA
# =========================

@app.route("/api/cash", methods=["GET", "POST"])
def api_cash():
    global cash_state

    if request.method == "GET":
        return jsonify(cash_state)

    cash_state = request.json
    return jsonify({"ok": True})


# =========================
# 📦 EXPORTAR COMPRAS EXCEL
# =========================

@app.route("/api/export/purchases/excel")
def export_purchases_excel():
    wb = Workbook()
    ws = wb.active

    ws.append(["Artigo", "Quantidade"])

    for p in purchases:
        ws.append([
            p.get("item"),
            p.get("qty")
        ])

    mem = BytesIO()
    wb.save(mem)
    mem.seek(0)

    return send_file(
        mem,
        as_attachment=True,
        download_name="compras.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# 📄 EXPORTAR COMPRAS PDF
# =========================

@app.route("/api/export/purchases/pdf")
def export_purchases_pdf():
    mem = BytesIO()
    c = canvas.Canvas(mem, pagesize=A4)

    y = 800
    c.drawString(50, y, "ARTIGOS A COMPRAR")
    y -= 30

    for p in purchases:
        linha = f"{p.get('item')} - {p.get('qty')}"
        c.drawString(50, y, linha)
        y -= 20

    c.save()
    mem.seek(0)

    return send_file(mem, as_attachment=True, download_name="compras.pdf")


# =========================
# 💰 EXPORTAR CAIXA EXCEL
# =========================

@app.route("/api/export/cash/excel")
def export_cash_excel():
    wb = Workbook()
    ws = wb.active

    ws.append(["Secção", "Fundo", "Entradas", "Saídas"])

    for sec in cash_state:
        s = cash_state[sec]
        ws.append([sec, s["start"], s["in"], s["out"]])

    mem = BytesIO()
    wb.save(mem)
    mem.seek(0)

    return send_file(mem, as_attachment=True, download_name="caixa.xlsx")


# =========================
# 📄 EXPORTAR CAIXA PDF
# =========================

@app.route("/api/export/cash/pdf")
def export_cash_pdf():
    mem = BytesIO()
    c = canvas.Canvas(mem, pagesize=A4)

    y = 800
    c.drawString(50, y, "FECHO DE CAIXA")
    y -= 30

    for sec in cash_state:
        s = cash_state[sec]
        c.drawString(50, y, f"{sec.upper()}")
        y -= 20
        c.drawString(50, y, f"Fundo: {s['start']} | Entradas: {s['in']} | Saídas: {s['out']}")
        y -= 30

    c.save()
    mem.seek(0)

    return send_file(mem, as_attachment=True, download_name="caixa.pdf")


# =========================
# 🚀 START
# =========================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)