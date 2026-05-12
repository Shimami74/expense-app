from flask import Flask, render_template, request, send_file
import sqlite3
import csv
from pathlib import Path

NAME_DICTIONARY_FILE = "name_dictionary.csv"

import os

from parser import process_pdf

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/download/<filename>")
def download_file(filename):

    return send_file(
        Path(OUTPUT_FOLDER) / filename,
        as_attachment=True
    )

@app.route("/upload", methods=["POST"])
def upload():

    file = request.files["pdf"]

    if file.filename == "":
        return "ファイル未選択"

    pdf_path = Path(UPLOAD_FOLDER) / file.filename

    file.save(pdf_path)

    output_excel = Path(OUTPUT_FOLDER) / "result.xlsx"

    result = process_pdf(pdf_path, output_excel)

    return render_template(
        "index.html",
        result={
            "total_rows": result["total_rows"],
            "unknown_count": result["unknown_count"],
            "total_amount": f"{result['total_amount']:,}",
            "download_file": "result.xlsx"
        }
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
        SELECT
            name,
            debit,
            description
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

        with open(
            NAME_DICTIONARY_FILE,
            "a",
            newline="",
            encoding="utf-8-sig"
        ) as f:

            writer = csv.writer(f)

            # 初回ヘッダ
            if not file_exists:
                writer.writerow(["alias", "real_name"])

            writer.writerow([alias, real_name])

    return redirect("/dictionary")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)