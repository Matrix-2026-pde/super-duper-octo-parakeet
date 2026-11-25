"""阶段二：SQLite 存储与基础 SQL。

核心流程：
    阶段一有效订单（DataFrame） -> 转为 SQLite 能接受的原生类型 -> 参数化 INSERT（UPSERT）
    -> 查询 / 汇总 / 修改 / 删除

设计要点：
- 使用 Python 标准库 sqlite3，无需额外安装依赖。
- 数据库列名使用英文，中文表头到英文字段的映射集中在 FIELD_MAP，便于 SQL 书写。
- 所有含外部数据的 SQL 都用 ? 占位符传参，不做字符串拼接（防止 SQL 注入）。
- 约束写在数据库内（CHECK / PRIMARY KEY / NOT NULL），即使绕过 Python 校验，数据库也会拒绝非法数据。

用法（项目根目录）：
    python -m app.db data/samples data/output/orders.db
"""

from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

import pandas as pd

from app.excel_batch import run_pipeline

# 建表语句。IF NOT EXISTS 让多次运行安全（幂等）。
SCHEMA = """
CREATE TABLE IF NOT EXISTS orders (
    order_id    TEXT    PRIMARY KEY,
    customer    TEXT    NOT NULL,
    product     TEXT    NOT NULL,
    quantity    INTEGER NOT NULL CHECK (quantity > 0),
    unit_price  REAL    NOT NULL CHECK (unit_price > 0),
    amount      REAL    NOT NULL,
    order_date  TEXT    NOT NULL,
    source_file TEXT,
    source_row  INTEGER
);
CREATE INDEX IF NOT EXISTS idx_orders_customer ON orders (customer);
CREATE INDEX IF NOT EXISTS idx_orders_date ON orders (order_date);
"""

# 阶段一 DataFrame 列名 -> 数据库字段名（插入时的顺序必须与 INSERT_SQL 一致）
FIELD_MAP = [
    ("订单号", "order_id"),
    ("客户", "customer"),
    ("产品", "product"),
    ("数量", "quantity"),
    ("单价", "unit_price"),
    ("金额", "amount"),
    ("日期", "order_date"),
    ("source_file", "source_file"),
    ("source_row", "source_row"),
]

# UPSERT：主键不存在则插入；主键已存在则用新值更新（excluded 指“本次准备插入的那一行”）
INSERT_SQL = """
INSERT INTO orders (order_id, customer, product, quantity, unit_price,
                    amount, order_date, source_file, source_row)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
ON CONFLICT (order_id) DO UPDATE SET
    customer    = excluded.customer,
    product     = excluded.product,
    quantity    = excluded.quantity,
    unit_price  = excluded.unit_price,
    amount      = excluded.amount,
    order_date  = excluded.order_date,
    source_file = excluded.source_file,
    source_row  = excluded.source_row
"""


# ---------------------------------------------------------------------------
# 1. 连接与建表
# ---------------------------------------------------------------------------
def connect(db_path) -> sqlite3.Connection:
    """打开数据库连接。db_path 可以是文件路径，也可以是 ':memory:'（内存库，测试用）。"""
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row  # 让查询结果可以按列名取值：row["customer"]
    return conn


def init_db(db_path) -> sqlite3.Connection:
    """连接数据库并建表（不存在时才创建）。"""
    conn = connect(db_path)
    conn.executescript(SCHEMA)
    return conn


# ---------------------------------------------------------------------------
# 2. 导入（写入）
# ---------------------------------------------------------------------------
def to_records(df: pd.DataFrame) -> list[tuple]:
    """把 DataFrame 行转换为 sqlite3 可接受的原生 Python 类型。

    sqlite3 不认识 numpy.int64、pandas.Timestamp 等类型，必须先转换；
    日期统一存为 'YYYY-MM-DD' 文本，便于比较和排序。
    """
    records = []
    for _, row in df.iterrows():
        records.append((
            str(row["订单号"]),
            str(row["客户"]),
            str(row["产品"]),
            int(row["数量"]),
            float(row["单价"]),
            float(row["金额"]),
            pd.Timestamp(row["日期"]).strftime("%Y-%m-%d"),
            str(row["source_file"]),
            int(row["source_row"]),
        ))
    return records


def import_orders(conn: sqlite3.Connection, df: pd.DataFrame) -> int:
    """批量写入订单，返回本次提交的行数。

    `with conn:` 是事务：全部成功则提交（commit），中途出错则整体回滚（rollback），
    不会出现“写了一半”的数据。
    """
    records = to_records(df)
    with conn:
        conn.executemany(INSERT_SQL, records)
    return len(records)


# ---------------------------------------------------------------------------
# 3. 查询与汇总（SELECT）
# ---------------------------------------------------------------------------
def get_order(conn: sqlite3.Connection, order_id: str) -> dict | None:
    """按主键查询单条订单；找不到返回 None。"""
    row = conn.execute(
        "SELECT * FROM orders WHERE order_id = ?", (order_id,)
    ).fetchone()
    return dict(row) if row else None


def list_orders(conn: sqlite3.Connection, customer: str | None = None) -> list[dict]:
    """列出订单；传入 customer 时只看该客户的订单。按日期、订单号排序。"""
    if customer is None:
        sql = "SELECT * FROM orders ORDER BY order_date, order_id"
        rows = conn.execute(sql).fetchall()
    else:
        sql = "SELECT * FROM orders WHERE customer = ? ORDER BY order_date, order_id"
        rows = conn.execute(sql, (customer,)).fetchall()
    return [dict(r) for r in rows]


def summarize_by_customer(conn: sqlite3.Connection) -> list[dict]:
    """按客户汇总：订单数、总数量、总金额，按总金额从大到小排列。"""
    sql = """
        SELECT customer,
               COUNT(*)                    AS order_count,
               SUM(quantity)               AS total_quantity,
               ROUND(SUM(amount), 2)       AS total_amount
        FROM orders
        GROUP BY customer
        ORDER BY total_amount DESC
    """
    return [dict(r) for r in conn.execute(sql).fetchall()]


# ---------------------------------------------------------------------------
# 4. 修改与删除（UPDATE / DELETE）
# ---------------------------------------------------------------------------
def update_quantity(conn: sqlite3.Connection, order_id: str, quantity: int) -> bool:
    """修改数量，同时按 金额 = 数量 × 单价 重新计算金额。返回是否找到并更新了该订单。

    金额是派生字段，修改数量时必须同步更新，否则报表会自相矛盾。
    数量非正数会被 CHECK 约束拒绝，抛出 sqlite3.IntegrityError。
    """
    sql = """
        UPDATE orders
        SET quantity = ?,
            amount   = ? * unit_price
        WHERE order_id = ?
    """
    with conn:
        cur = conn.execute(sql, (quantity, quantity, order_id))
    return cur.rowcount > 0


def delete_order(conn: sqlite3.Connection, order_id: str) -> bool:
    """删除订单，返回是否真的删除了一行。"""
    with conn:
        cur = conn.execute("DELETE FROM orders WHERE order_id = ?", (order_id,))
    return cur.rowcount > 0


# ---------------------------------------------------------------------------
# 命令行入口：流水线 -> 写入数据库
# ---------------------------------------------------------------------------
def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print("用法: python -m app.db <输入目录> <数据库文件.db>")
        return 1

    kept, _rejected, stats = run_pipeline(Path(argv[1]))
    conn = init_db(argv[2])
    written = import_orders(conn, kept)
    total = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    conn.close()

    for key, value in stats.items():
        print(f"{key}: {value}")
    print(f"本次写入（新增或更新）: {written}")
    print(f"数据库当前订单总数: {total}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
