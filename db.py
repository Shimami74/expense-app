import sqlite3
from pathlib import Path

DB_PATH = Path("expense.db")


def init_db():

    conn = sqlite3.connect(DB_PATH)

    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS expenses (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        page INTEGER,

        slip_no TEXT,

        name TEXT,

        debit INTEGER,

        description TEXT

    )
    """)

    conn.commit()

    conn.close()


def save_to_db(df):

    conn = sqlite3.connect(DB_PATH)

    cur = conn.cursor()

    for _, row in df.iterrows():

        cur.execute("""
        INSERT INTO expenses (
            page,
            slip_no,
            name,
            debit,
            description
        )
        VALUES (?, ?, ?, ?, ?)
        """, (

            row.get("ページ"),
            row.get("伝票番号"),
            row.get("氏名"),
            row.get("借方金額"),
            row.get("摘要")

        ))

    conn.commit()

    conn.close()