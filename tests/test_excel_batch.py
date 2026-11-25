"""阶段一测试：清洗、校验、去重、端到端流程。"""

import pandas as pd
import pytest

from app.excel_batch import (
    clean_orders,
    deduplicate_orders,
    process_directory,
    validate_orders,
)

COLUMNS = ["订单号", "客户", "产品", "数量", "单价", "日期"]


def _raw(rows):
    """构造与 read_order_files 输出结构一致的原始 DataFrame。"""
    df = pd.DataFrame(rows, columns=COLUMNS)
    df["source_file"] = "test.xlsx"
    df["source_row"] = range(2, 2 + len(df))
    return df


# ---------------------------------------------------------------------------
# 清洗
# ---------------------------------------------------------------------------
def test_clean_strips_whitespace_and_converts_types():
    raw = _raw([["SO-1", "  华北  ", "显示器", "2", "999.5", "2026-01-05"]])
    out = clean_orders(raw)
    assert out.loc[0, "客户"] == "华北"
    assert out.loc[0, "数量"] == 2
    assert out.loc[0, "单价"] == pytest.approx(999.5)
    assert out.loc[0, "金额"] == pytest.approx(1999.0)
    assert out.loc[0, "日期"] == pd.Timestamp("2026-01-05")


def test_clean_turns_blank_text_into_none():
    raw = _raw([["SO-1", "   ", "显示器", 1, 10, "2026-01-05"]])
    out = clean_orders(raw)
    assert out.loc[0, "客户"] is None


# ---------------------------------------------------------------------------
# 校验
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "row, expected_error",
    [
        (["", "客户A", "产品", 1, 10, "2026-01-01"], "订单号为空"),
        (["SO-1", "", "产品", 1, 10, "2026-01-01"], "客户为空"),
        (["SO-1", "客户A", "产品", 0, 10, "2026-01-01"], "数量非正数或非数字"),
        (["SO-1", "客户A", "产品", 1, -5, "2026-01-01"], "单价非正数或非数字"),
        (["SO-1", "客户A", "产品", 1, 10, "bad-date"], "日期无法解析"),
    ],
)
def test_validate_rejects_bad_rows(row, expected_error):
    cleaned = clean_orders(_raw([row]))
    valid, invalid = validate_orders(cleaned)
    assert valid.empty
    assert expected_error in invalid.iloc[0]["error"]


def test_validate_keeps_good_rows():
    cleaned = clean_orders(_raw([["SO-1", "客户A", "产品", 1, 10, "2026-01-01"]]))
    valid, invalid = validate_orders(cleaned)
    assert len(valid) == 1
    assert invalid.empty


# ---------------------------------------------------------------------------
# 去重
# ---------------------------------------------------------------------------
def test_deduplicate_keeps_first_occurrence():
    cleaned = clean_orders(
        _raw(
            [
                ["SO-1", "A", "p", 1, 10, "2026-01-01"],
                ["SO-1", "B", "p", 2, 10, "2026-01-02"],
            ]
        )
    )
    kept, dups = deduplicate_orders(cleaned)
    assert len(kept) == 1
    assert kept.iloc[0]["客户"] == "A"
    assert dups.iloc[0]["error"] == "订单号重复"


# ---------------------------------------------------------------------------
# 端到端
# ---------------------------------------------------------------------------
def test_process_directory_end_to_end(tmp_path):
    input_dir = tmp_path / "in"
    input_dir.mkdir()
    _raw(
        [
            ["SO-1", "A", "p", 1, 10, "2026-01-01"],
            ["SO-2", "B", "p", 0, 10, "2026-01-01"],
            ["SO-1", "A", "p", 1, 10, "2026-01-01"],
        ]
    ).drop(columns=["source_file", "source_row"]).to_excel(input_dir / "a.xlsx", index=False)

    output = tmp_path / "out" / "report.xlsx"
    stats = process_directory(input_dir, output)

    assert stats["读取行数"] == 3
    assert stats["有效行数"] == 1
    assert stats["校验失败"] == 1
    assert stats["重复订单"] == 1
    assert output.exists()

    detail = pd.read_excel(output, sheet_name="订单明细")
    reject = pd.read_excel(output, sheet_name="异常记录")
    assert list(detail["订单号"]) == ["SO-1"]
    assert len(reject) == 2
    assert set(reject["source_file"]) == {"a.xlsx"}
