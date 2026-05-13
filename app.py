from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import sqlite3
import csv
import uuid

from pathlib import Path
from datetime import datetime

from parser import process_pdf


# =========================
# 名前辞書読込
# =========================
def load_dictionary():

    d = {}

    try:
        with open("name_dictionary.csv", encoding="utf-8-sig") as f:

            reader = csv.DictReader(f)

            for row in reader:
                d[row["alias"]] = row["real_name"]

    except:
        pass

    return d


# =========================
# Flask
# =========================
app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

NAME_DICTIONARY_FILE = "name_dictionary.csv"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# =========================
# TOP
# =========================
@app.route("/")
def index():

    return render_template("index.html")


# =========================
# PDFアップロード
# =========================
@app.route("/upload", methods=["POST"])
def upload():

    file = request.files.get("pdf")

    if not file or file.filename == "":
        return redirect(url_for("index"))

    # =========================
    # 元ファイル名
    # =========================
    original_name = Path(file.filename).stem

    # =========================
    # PDF保存名（内部用）
    # =========================
    safe_name = f"{uuid.uuid4()}.pdf"

    pdf_path = Path(UPLOAD_FOLDER) / safe_name

    file.save(pdf_path)

    # =========================
    # 出力Excel名
    # =========================
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    excel_name = f"{original_name}_{timestamp}.xlsx"

    output_excel = Path(OUTPUT_FOLDER) / excel_name

    # =========================
    # PDF解析
    # =========================
    result = process_pdf(pdf_path, output_excel)

    # PDF削除
    if pdf_path.exists():
        os.remove(pdf_path)

    print("DEBUG result:", result)

    # =========================
    # 結果表示
    # =========================
    return render_template(
        "index.html",
        result={
            "total_rows": result["total_rows"],
            "unknown_count": result["unknown_count"],
            "total_amount": f"{result['total_amount']:,}",
            "download_file": output_excel.name
        }
    )


# =========================
# ダウンロード
# =========================
@app.route("/download/<filename>")
def download_file(filename):

    return send_file(
        Path(OUTPUT_FOLDER) / filename,
        as_attachment=True,
        download_name=filename
    )


# =========================
# 検索
# =========================
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


# =========================
# 辞書画面
# =========================
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


# =========================
# 辞書追加
# =========================
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

            if not file_exists:
                writer.writerow(["alias", "real_name"])

            writer.writerow([alias, real_name])

    return redirect(url_for("dictionary"))


# =========================
# 起動
# =========================
if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )

