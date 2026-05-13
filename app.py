from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import sqlite3
import csv
from pathlib import Path
import uuid

from parser import process_pdf

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
NAME_DICTIONARY_FILE = "name_dictionary.csv"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")
def index():

    # フラッシュ的に結果表示したい場合はここに拡張可能
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():

    file = request.files.get("pdf")

    if not file or file.filename == "":
        return redirect(url_for("index"))

    # 安全なファイル名
    safe_name = f"{uuid.uuid4()}.pdf"
    pdf_path = Path(UPLOAD_FOLDER) / safe_name

    file.save(pdf_path)

    # 出力もユニーク化（重要）
    output_excel = Path(OUTPUT_FOLDER) / f"{uuid.uuid4()}.xlsx"

    result = process_pdf(pdf_path, output_excel)

    from openpyxl import load_workbook

    wb = load_workbook(output_excel)
    ws = wb["抽出データ_氏名あり"]

    ws.auto_filter.ref = ws.dimensions

    wb.save(output_excel)
    wb.close()

    # PRGパターン（重要）
    return render_template(
        "index.html",
        result={
            "total_rows": result["total_rows"],
            "unknown_count": result["unknown_count"],
            "total_amount": f"{result['total_amount']:,}",
            "download_file": output_excel.name
        }
    )


@app.route("/download/<filename>")
def download_file(filename):
    return send_file(
        Path(OUTPUT_FOLDER) / filename,
        as_attachment=True
    )


@app.route("/search")
def search():

    keyword = request.args.get("keyword", "")
    results = []

    if keyword:

        conn = sqlite3.connect("expense.db")
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        cur.execute("""
            SELECT name, debit, description
            FROM expenses
            WHERE name LIKE ?
            ORDER BY debit DESC
        """, (f"%{keyword}%",))

        results = cur.fetchall()
        conn.close()

    return render_template(
        "search.html",
        results=results,
        keyword=keyword
    )


@app.route("/dictionary")
def dictionary():

    rows = []
    path = Path(NAME_DICTIONARY_FILE)

    if path.exists():

        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            rows = list(reader)

    return render_template(
        "dictionary.html",
        rows=rows
    )


@app.route("/dictionary/add", methods=["POST"])
def add_dictionary():

    alias = request.form.get("alias", "").strip()
    real_name = request.form.get("real_name", "").strip()

    if alias and real_name:

        file_exists = Path(NAME_DICTIONARY_FILE).exists()

        with open(NAME_DICTIONARY_FILE, "a", newline="", encoding="utf-8-sig") as f:

            writer = csv.writer(f)

            if not file_exists:
                writer.writerow(["alias", "real_name"])

            writer.writerow([alias, real_name])

    return redirect(url_for("dictionary"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)