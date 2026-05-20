from flask import Flask, request, jsonify, render_template
from dotenv import load_dotenv
from datetime import datetime
import requests
import os
from flask import send_file
from io import BytesIO
from openpyxl import Workbook

load_dotenv()

app = Flask(__name__)

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

TABELA = "leituras_ruido"

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation"
}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/ruido", methods=["POST"])
def receber_ruido():
    try:
        dados = request.get_json()

        if not dados:
            return jsonify({"status": "erro", "mensagem": "JSON vazio ou inválido"}), 400

        leitura = {
            "operador": str(dados.get("operador", "1")),
            "decibeis": float(dados.get("decibeis", 0))
        }

        url = f"{SUPABASE_URL}/rest/v1/{TABELA}"

        resposta = requests.post(url, json=leitura, headers=HEADERS)

        if resposta.status_code not in [200, 201]:
            return jsonify({
                "status": "erro",
                "mensagem": resposta.text
            }), 500

        return jsonify({
            "status": "ok",
            "leitura": resposta.json()[0]
        }), 200

    except Exception as erro:
        return jsonify({
            "status": "erro",
            "mensagem": str(erro)
        }), 500


@app.route("/api/ultimas", methods=["GET"])
def ultimas():
    try:
        url = f"{SUPABASE_URL}/rest/v1/{TABELA}"

        params = {
            "select": "*",
            "order": "data_hora.desc",
            "limit": "30"
        }

        resposta = requests.get(url, headers=HEADERS, params=params)

        if resposta.status_code != 200:
            return jsonify({
                "status": "erro",
                "mensagem": resposta.text
            }), 500

        dados = resposta.json()

        leituras = []

        for item in dados:
            leituras.append({
                "operador": item["operador"],
                "decibeis": float(item["decibeis"]),
                "data_hora": formatar_data(item["data_hora"])
            })

        return jsonify(leituras), 200

    except Exception as erro:
        return jsonify({
            "status": "erro",
            "mensagem": str(erro)
        }), 500


def formatar_data(data_supabase):
    try:
        data = datetime.fromisoformat(data_supabase.replace("Z", "+00:00"))
        return data.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return data_supabase

@app.route("/relatorio_excel", methods=["GET"])
def relatorio_excel():
    try:
        url = f"{SUPABASE_URL}/rest/v1/{TABELA}"

        params = {
            "select": "*",
            "order": "data_hora.desc"
        }

        resposta = requests.get(url, headers=HEADERS, params=params)

        if resposta.status_code != 200:
            return jsonify({
                "status": "erro",
                "mensagem": resposta.text
            }), 500

        dados = resposta.json()

        wb = Workbook()
        ws = wb.active
        ws.title = "Leituras de Ruído"

        ws.append(["ID", "Operador", "Decibéis", "Data e Hora", "Status"])

        for item in dados:
            decibeis = float(item["decibeis"])

            if decibeis >= 95:
                status = "Crítico"
            elif decibeis >= 90:
                status = "Alerta"
            elif decibeis >= 85:
                status = "Atenção"
            else:
                status = "Normal"

            ws.append([
                item.get("id", ""),
                item.get("operador", ""),
                decibeis,
                formatar_data(item.get("data_hora", "")),
                status
            ])

        arquivo = BytesIO()
        wb.save(arquivo)
        arquivo.seek(0)

        return send_file(
            arquivo,
            as_attachment=True,
            download_name="relatorio_ruido.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as erro:
        return jsonify({
            "status": "erro",
            "mensagem": str(erro)
        }), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)