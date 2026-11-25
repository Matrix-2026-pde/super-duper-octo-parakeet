"""阶段二测试：建表、导入（含 UPSERT）、查询、汇总、修改、删除、约束。

全部使用内存数据库 ':memory:'，不产生磁盘文件，每个测试互相独立。
"""

import sqlite3

import pandas as pd
import pytest

from app.db import (
    delete_order,
    get_order,
    import_orders,
    init_db,
    list_orders,
    summarize_by_customer,
    update_quantity,
)


def _df(rows):
    """构造与阶段一输出结构一致的 DataFrame（有效订单）。"""
    df = pd.DataFrame(
        rows,
        columns=["订单号", "客户", "产品", "数量", "单价", "金额", "日期", "source_file", "source_row"],
    )
    df["日期"] = pd.to_datetime(df["日期"])
    return df


@pytest.fixture
def conn():
    c = init_db(":memory:")
    yield c
    c.close()


@pytest.fixture
def seeded(conn):
    """预先写入三条订单的数据库连接。"""
    import_orders(conn, _df([
        ["SO-1", "华北贸易", "显示器", 2, 999.0, 1998.0, "2026-01-05", "a.xlsx", 2],
        ["SO-2", "东方科技", "键盘", 10, 129.5, 1295.0, "2026-01-07", "a.xlsx", 3],
        ["SO-3", "华北贸易", "鼠标", 3, 59.0, 177.0, "2026-02-01", "b.xlsx", 2],
    ]))
    return conn


# ---------------------------------------------------------------------------
# 建表与导入
# ---------------------------------------------------------------------------
def test_init_db_is_idempotent(conn):
    # 重复执行建表语句不应报错（IF NOT EXISTS）
    conn.executescript(
        "CREATE TABLE IF NOT EXISTS orders (order_id TEXT PRIMARY KEY);"
    )
    tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")]
    assert "orders" in tables


def test_import_writes_all_rows(conn):
    written = import_orders(conn, _df([
        ["SO-1", "A", "p", 1, 10.0, 10.0, "2026-01-01", "a.xlsx", 2],
        ["SO-2", "B", "p", 2, 10.0, 20.0, "2026-01-02", "a.xlsx", 3],
    ]))
    assert written == 2
    assert conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0] == 2


def test_import_upsert_updates_existing_order(seeded):
    # 同一订单号再次导入：应更新而不是报主键冲突，且总数不变
    import_orders(seeded, _df([
        ["SO-1", "华北贸易", "显示器", 5, 999.0, 4995.0, "2026-01-05", "c.xlsx", 9],
    ]))
    row = get_order(seeded, "SO-1")
    assert row["quantity"] == 5
    assert row["amount"] == pytest.approx(4995.0)
    assert row["source_file"] == "c.xlsx"
    assert seeded.execute("SELECT COUNT(*) FROM orders").fetchone()[0] == 3


def test_import_is_atomic_on_failure(conn):
    # 第二行数量为 0 违反 CHECK 约束，整批应全部回滚，不留半截数据
    bad = _df([
        ["SO-1", "A", "p", 1, 10.0, 10.0, "2026-01-01", "a.xlsx", 2],
        ["SO-2", "B", "p", 0, 10.0, 0.0, "2026-01-02", "a.xlsx", 3],
    ])
    with pytest.raises(sqlite3.IntegrityError):
        import_orders(conn, bad)
    assert conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0] == 0


def test_database_rejects_invalid_quantity_directly(conn):
    # 即使绕过 Python 校验直接写 SQL，数据库约束也会拒绝非法数量
    with pytest.raises(sqlite3.IntegrityError):
        conn.execute(
            "INSERT INTO orders VALUES ('X', 'a', 'b', 0, 10, 0, '2026-01-01', 'f', 1)"
        )


# ---------------------------------------------------------------------------
# 查询与汇总
# ---------------------------------------------------------------------------
def test_get_order_found_and_missing(seeded):
    assert get_order(seeded, "SO-2")["customer"] == "东方科技"
    assert get_order(seeded, "NOPE") is None


def test_list_orders_filter_by_customer(seeded):
    rows = list_orders(seeded, customer="华北贸易")
    assert [r["order_id"] for r in rows] == ["SO-1", "SO-3"]


def test_list_orders_all_sorted_by_date(seeded):
    dates = [r["order_date"] for r in list_orders(seeded)]
    assert dates == sorted(dates)


def test_summarize_by_customer(seeded):
    summary = summarize_by_customer(seeded)
    # 按总金额降序：华北贸易 1998+177=2175 > 东方科技 1295
    assert summary[0]["customer"] == "华北贸易"
    assert summary[0]["order_count"] == 2
    assert summary[0]["total_quantity"] == 5
    assert summary[0]["total_amount"] == pytest.approx(2175.0)


# ---------------------------------------------------------------------------
# 修改与删除
# ---------------------------------------------------------------------------
def test_update_quantity_recalculates_amount(seeded):
    assert update_quantity(seeded, "SO-1", 4) is True
    row = get_order(seeded, "SO-1")
    assert row["quantity"] == 4
    assert row["amount"] == pytest.approx(4 * 999.0)


def test_update_quantity_missing_order_returns_false(seeded):
    assert update_quantity(seeded, "NOPE", 4) is False


def test_update_quantity_rejects_non_positive(seeded):
    with pytest.raises(sqlite3.IntegrityError):
        update_quantity(seeded, "SO-1", 0)
    # 失败后原数据保持不变
    assert get_order(seeded, "SO-1")["quantity"] == 2


def test_delete_order(seeded):
    assert delete_order(seeded, "SO-2") is True
    assert get_order(seeded, "SO-2") is None
    assert delete_order(seeded, "SO-2") is False
