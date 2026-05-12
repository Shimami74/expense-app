from pathlib import Path
import re
import csv
import fitz  # PyMuPDF
import pandas as pd

from db import init_db, save_to_db

from openpyxl.utils import get_column_letter

# Janomeを使う場合
# pip install janome
try:
    from janome.tokenizer import Tokenizer

    tokenizer = Tokenizer()
    JANOME_AVAILABLE = True

except:
    JANOME_AVAILABLE = False


# =========================
# 設定
# =========================
DEBUG_PARSED = "debug_parsed_entries.txt"

NAME_DICTIONARY_FILE = "name_dictionary.csv"


# =========================
# ユーティリティ
# =========================
def clean_text(s: str) -> str:

    if s is None:
        return ""

    s = str(s)

    s = s.replace("\u3000", " ")
    s = s.replace("\xa0", " ")
    s = s.replace("\r", "")

    return s.strip()


def normalize_spaces(s: str) -> str:
    return re.sub(r"[ \t]+", " ", clean_text(s))


def load_name_dictionary():

    result = {}

    path = Path(NAME_DICTIONARY_FILE)

    if not path.exists():
        return result

    with open(path, "r", encoding="utf-8-sig") as f:

        reader = csv.DictReader(f)

        for row in reader:

            alias = clean_text(row.get("alias"))
            real_name = clean_text(row.get("real_name"))

            if alias and real_name:

                alias = alias.strip()
                real_name = real_name.strip()

                result[alias] = real_name

    return result


# 学習辞書読込
CUSTOM_NAME_DICT = load_name_dictionary()


def to_int_amount(s: str) -> int:
    return int(str(s).replace(",", "").strip())


def normalize_name(name: str) -> str:

    if not name:
        return ""

    name = clean_text(name)

    # 不要文字除去
    name = re.sub(r"[■□№]", "", name)

    # 全角スペース除去
    name = name.replace("\u3000", "")

    name = name.strip()

    # 学習辞書
    if name in CUSTOM_NAME_DICT:
        return CUSTOM_NAME_DICT[name]

    return name


# =========================
# PDFテキスト取得
# =========================
def extract_page_lines(pdf_path: Path):

    pages = []

    doc = fitz.open(str(pdf_path))

    try:
        for page_no in range(len(doc)):

            page = doc.load_page(page_no)

            text = page.get_text("text", sort=True) or ""

            lines = [
                clean_text(line)
                for line in text.splitlines()
            ]

            # 空行除去
            lines = [line for line in lines if line]

            pages.append((page_no + 1, lines))

    finally:
        doc.close()

    return pages


# =========================
# 明細開始判定
# =========================
def is_entry_start_line(line: str) -> bool:

    line = normalize_spaces(line)

    return bool(
        re.match(
            r"^\d{4}/\d{2}/\d{2}\s+\d{4}/\d{2}/\d{2}\s+83330\s+旅費\b",
            line
        )
    )


# =========================
# 金額抽出
# =========================
def extract_amounts_from_line1(line1: str):

    s = normalize_spaces(line1)

    nums = re.findall(r"\d{1,3}(?:,\d{3})*", s)

    if len(nums) < 2:
        return None, None

    debit = to_int_amount(nums[-2])
    balance = to_int_amount(nums[-1])

    return debit, balance


# =========================
# 伝票番号等抽出
# =========================
def extract_slip_and_subcode_from_line2(line2: str):

    s = normalize_spaces(line2)

    m = re.match(
        r"^(?P<slip>G\d+)\s+(?P<subcode>7101[23])\s+(?P<subname>.+)$",
        s
    )

    if not m:
        return "", "", ""

    return (
        m.group("slip"),
        m.group("subcode"),
        m.group("subname")
    )


# =========================
# 部門情報抽出
# =========================
def extract_line_no_and_dept_from_line3(line3: str):

    s = normalize_spaces(line3)

    m = re.match(
        r"^(?P<line_no>\d+)\s+(?P<tax>\S+)\s+(?P<dept_code>\d+)\s+(?P<dept_name>.+)$",
        s
    )

    if not m:
        return "", "", "", ""

    return (
        m.group("line_no"),
        m.group("tax"),
        m.group("dept_code"),
        m.group("dept_name")
    )


# =========================
# 摘要
# =========================
def extract_description_from_line4(line4: str):
    return clean_text(line4)


# =========================
# 氏名抽出
# =========================

COMMON_NON_NAMES = {
    "東京", "大阪", "会議", "打合", "レンタカー",
    "高速代", "宿泊", "技術", "建設", "補修",
    "タクシー", "新幹線", "立替", "精算",
    "駐車場", "ガソリン", "羽田", "伊丹",
    "品川", "名古屋", "福岡"
}

NAME_PATTERNS = [

    # ■山本健太
    {
        "pattern": r"■([一-龥々ぁ-んァ-ヶ]{2,8})",
        "score": 10
    },

    # 山本健太2/2№46593
    {
        "pattern": r"([一-龥々ぁ-んァ-ヶ]{2,8})(?=\d{1,2}/\d{1,2})",
        "score": 8
    },

    # 山本健太 03/02
    {
        "pattern": r"([一-龥々ぁ-んァ-ヶ]{2,8})(?=\s+\d{1,2}/\d{1,2})",
        "score": 7
    },

    # 山本健太MF49130
    {
        "pattern": r"([一-龥々ぁ-んァ-ヶ]{2,8})(?=\s*MF\d+)",
        "score": 8
    },

    # 山本健太 41725000-20
    {
        "pattern": r"([一-龥々ぁ-んァ-ヶ]{2,8})(?=\s+\d{5,}(?:-\d+)?)",
        "score": 6
    },

    # 行末
    {
        "pattern": r"(?:\s|　)([一-龥々ぁ-んァ-ヶ]{2,8})$",
        "score": 3
    },
]


def score_name(name: str) -> int:

    if not name:
        return -999

    if name in COMMON_NON_NAMES:
        return -999

    score = 0

    # フルネームっぽい
    if 4 <= len(name) <= 6:
        score += 5

    # 姓のみ
    elif 2 <= len(name) <= 3:
        score += 2

    # 数字混じり除外
    if re.search(r"\d", name):
        score -= 10

    # 英字混じり除外
    if re.search(r"[A-Za-z]", name):
        score -= 10

    return score


def extract_name(desc: str) -> str:

    desc = clean_text(desc)

    candidates = []

    # =====================
    # 正規表現抽出
    # =====================
    for item in NAME_PATTERNS:

        pattern = item["pattern"]
        base_score = item["score"]

        matches = re.finditer(pattern, desc)

        for m in matches:

            try:
                name = normalize_name(m.group(1))

            except:
                continue

            score = base_score + score_name(name)

            candidates.append({
                "name": name,
                "score": score
            })

    # スコア順
    if candidates:

        candidates.sort(
            key=lambda x: x["score"],
            reverse=True
        )

        best = candidates[0]["name"]

        return best

    # =====================
    # Janome fallback
    # =====================
    if JANOME_AVAILABLE:

        for token in tokenizer.tokenize(desc):

            pos = token.part_of_speech.split(",")

            if (
                pos[0] == "名詞"
                and pos[1] == "固有名詞"
                and pos[2] == "人名"
            ):

                name = normalize_name(token.surface)

                if score_name(name) > 0:
                    return name

    return ""


# =========================
# 明細抽出
# =========================
def parse_entries_from_pages(pages):

    rows = []
    debug_rows = []

    for page_no, lines in pages:

        i = 0

        while i < len(lines):

            line1 = lines[i]

            if not is_entry_start_line(line1):
                i += 1
                continue

            if i + 3 >= len(lines):
                break

            line2 = lines[i + 1]
            line3 = lines[i + 2]
            line4 = lines[i + 3]

            # G番号確認
            if not re.match(r"^G\d+", normalize_spaces(line2)):
                i += 1
                continue

            debit, balance = extract_amounts_from_line1(line1)

            slip_no, subcode, subname = (
                extract_slip_and_subcode_from_line2(line2)
            )

            line_no, tax, dept_code, dept_name = (
                extract_line_no_and_dept_from_line3(line3)
            )

            desc = extract_description_from_line4(line4)

            name = extract_name(desc)

            rows.append({
                "ページ": page_no,
                "計上行": normalize_spaces(line1),
                "伝票番号": slip_no,
                "相手補助科目コード": subcode,
                "相手補助科目": subname,
                "行": line_no,
                "税": tax,
                "部門コード": dept_code,
                "部門名": dept_name,
                "摘要": desc,
                "氏名": name,
                "借方金額": debit if debit is not None else None,
                "残高": balance if balance is not None else None,
            })

            debug_rows.append(
                f"[PAGE {page_no}] "
                f"伝票={slip_no}, "
                f"行={line_no}, "
                f"借方={debit}, "
                f"氏名={name}, "
                f"摘要={desc}"
            )

            i += 4

    with open(DEBUG_PARSED, "w", encoding="utf-8") as f:

        for row in debug_rows:
            f.write(row + "\n")

    return pd.DataFrame(rows)


# =========================
# Excel列幅調整
# =========================
def auto_adjust_columns(writer, sheet_name, df):

    ws = writer.sheets[sheet_name]

    for idx, col in enumerate(df.columns, start=1):

        max_len = len(str(col))

        for v in df[col].astype(str).fillna(""):

            max_len = max(
                max_len,
                len(str(v))
            )

        ws.column_dimensions[
            get_column_letter(idx)
        ].width = min(max_len + 2, 80)


# =========================
# メイン処理
# =========================
def process_pdf(pdf_path, output_excel):

    print("===================================")
    print("PDF読込開始")
    print("===================================")

    if not pdf_path.exists():

        print("PDFが見つかりません。")
        return

    pages = extract_page_lines(pdf_path)

    df = parse_entries_from_pages(pages)

    print("debug解析結果出力:", DEBUG_PARSED)

    if df.empty:

        print("取引データを抽出できませんでした。")
        return

    # 氏名あり/なし
    unknown_df = df[df["氏名"] == ""].copy()

    named_df = df[df["氏名"] != ""].copy()

    # 氏名別集計
    summary = (
        named_df
        .groupby("氏名", as_index=False)["借方金額"]
        .sum()
        .rename(columns={
            "借方金額": "合計借方金額"
        })
        .sort_values(
            "合計借方金額",
            ascending=False
        )
    )

    total_all_rows = df["借方金額"].fillna(0).sum()

    with pd.ExcelWriter(
        str(output_excel),
        engine="openpyxl"
    ) as writer:

        named_df.to_excel(
            writer,
            sheet_name="抽出データ_氏名あり",
            index=False
        )

        summary.to_excel(
            writer,
            sheet_name="氏名別集計",
            index=False
        )

        if not unknown_df.empty:

            unknown_df.to_excel(
                writer,
                sheet_name="氏名未抽出_要確認",
                index=False
            )

        auto_adjust_columns(
            writer,
            "抽出データ_氏名あり",
            named_df
        )

        auto_adjust_columns(
            writer,
            "氏名別集計",
            summary
        )

        if not unknown_df.empty:

            auto_adjust_columns(
                writer,
                "氏名未抽出_要確認",
                unknown_df
            )

        # 金額フォーマット
        for sheet_name in writer.sheets:

            ws = writer.sheets[sheet_name]

            headers = {
                ws.cell(row=1, column=col).value: col
                for col in range(1, ws.max_column + 1)
            }

            for money_col_name in [
                "借方金額",
                "合計借方金額",
                "残高"
            ]:

                if money_col_name in headers:

                    col_idx = headers[money_col_name]

                    for row in range(2, ws.max_row + 1):

                        ws.cell(
                            row=row,
                            column=col_idx
                        ).number_format = "#,##0"

    print("\n===================================")
    print("完了")
    print("===================================")

    print(f"Excelを出力しました: {output_excel}")
    print(f"抽出した取引件数: {len(df)}")
    print(f"借方合計: {total_all_rows:,}")

    if not unknown_df.empty:

        print(f"氏名未抽出件数: {len(unknown_df)}")
        print("『氏名未抽出_要確認』シートを確認してください。")

    else:
        print("すべての行で氏名を抽出できました。")

    print("\n=== 氏名別集計（上位20件） ===")

    print(
        summary.head(20).to_string(index=False)
    )

    # DB保存
    init_db()

    save_to_db(df)

    return {
        "output_excel": output_excel,
        "total_rows": len(df),
        "unknown_count": len(unknown_df),
        "total_amount": int(total_all_rows),
    }