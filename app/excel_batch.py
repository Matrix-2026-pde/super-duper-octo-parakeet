"""阶段一：Excel 订单批量处理。

流程：
    读取目录下所有 .xlsx -> 合并为一张表 -> 清洗 -> 校验 -> 去重 -> 导出 Excel 报表

设计要点：
- 清洗、校验、去重都是纯函数（输入 DataFrame，输出 DataFrame），便于 pytest 单独测试。
- 每条被拒绝的记录都保留 source_file 与 source_row，便于回溯到原 Excel 行。

用法（项目根目录）：
    python -m app.excel_batch data/samples data/output/report.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

# 标准列名（输入 Excel 表头必须与之一致）
REQUIRED_COLUMNS = ["订单号", "客户", "产品", "数量", "单价", "日期"]
# 记录来源信息的辅助列
SOURCE_COLUMNS = ["source_file", "source_row"]


# ---------------------------------------------------------------------------
# 1. 读取
# ---------------------------------------------------------------------------
def read_order_files(input_dir: Path) -> pd.DataFrame:
    """读取目录下所有 .xlsx 文件并合并为一张 DataFrame。

    每行附带 source_file（文件名）与 source_row（Excel 中的行号，从 1 计，含表头）。
    缺少必需列的文件会抛出 ValueError，避免静默丢数据。
    """
    input_dir = Path(input_dir)
    files = sorted(input_dir.glob("*.xlsx"))
    if not files:
        raise FileNotFoundError(f"目录中没有 .xlsx 文件: {input_dir}")

    frames = []
    for path in files:
        df = pd.read_excel(path, dtype=object)
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(f"{path.name} 缺少列: {missing}")

        df = df[REQUIRED_COLUMNS].copy()
        # 表头占第 1 行，数据从第 2 行开始；reset_index 后 index 0 对应第 2 行
        df["source_file"] = path.name
        df["source_row"] = df.index + 2
        frames.append(df)

    return pd.concat(frames, ignore_index=True)


# ---------------------------------------------------------------------------
# 2. 清洗
# ---------------------------------------------------------------------------
def clean_orders(df: pd.DataFrame) -> pd.DataFrame:
    """清洗：去除文本首尾空白、统一空字符串为 None、转换数量/单价/日期类型。

    无法转换的值保留为 None，交给 validate_orders 统一判定，不在这里丢弃。
    """
    out = df.copy()

    for col in ["订单号", "客户", "产品"]:
        out[col] = out[col].apply(_clean_text)

    out["数量"] = pd.to_numeric(out["数量"].apply(_clean_text), errors="coerce")
    out["单价"] = pd.to_numeric(out["单价"].apply(_clean_text), errors="coerce")
    out["日期"] = pd.to_datetime(out["日期"].apply(_clean_text), errors="coerce")

    # 金额 = 数量 × 单价，便于报表直接使用
    out["金额"] = (out["数量"] * out["单价"]).round(2)
    return out


def _clean_text(value):
    """文本清洗：None/NaN/空白 -> None，其余去首尾空白。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    return text if text else None


# ---------------------------------------------------------------------------
# 3. 校验
# ---------------------------------------------------------------------------
def validate_orders(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """校验：返回 (有效记录, 无效记录)。无效记录增加 error 列说明原因。

    规则：
    - 订单号、客户、产品不能为空
    - 数量、单价必须是大于 0 的数字
    - 日期必须可解析
    """
    messages = []
    for _, row in df.iterrows():
        row_errors = []
        if row["订单号"] is None:
            row_errors.append("订单号为空")
        if row["客户"] is None:
            row_errors.append("客户为空")
        if row["产品"] is None:
            row_errors.append("产品为空")
        if not _is_positive_number(row["数量"]):
            row_errors.append("数量非正数或非数字")
        if not _is_positive_number(row["单价"]):
            row_errors.append("单价非正数或非数字")
        if pd.isna(row["日期"]):
            row_errors.append("日期无法解析")
        messages.append("; ".join(row_errors))

    messages = pd.Series(messages, index=df.index)
    invalid_mask = messages != ""
    valid = df[~invalid_mask].copy()
    invalid = df[invalid_mask].copy()
    invalid["error"] = messages[invalid_mask]
    return valid, invalid


def _is_positive_number(value) -> bool:
    """判断单个值是否为大于 0 的数字（None/NaN/非数字均视为不合格）。"""
    return pd.notna(value) and value > 0


# ---------------------------------------------------------------------------
# 4. 去重
# ---------------------------------------------------------------------------
def deduplicate_orders(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """按订单号去重：保留第一次出现的记录，返回 (保留记录, 重复记录)。"""
    duplicated_mask = df.duplicated(subset=["订单号"], keep="first")
    kept = df[~duplicated_mask].copy()
    duplicates = df[duplicated_mask].copy()
    duplicates["error"] = "订单号重复"
    return kept, duplicates


# ---------------------------------------------------------------------------
# 5. 导出
# ---------------------------------------------------------------------------
def export_report(
    valid: pd.DataFrame,
    rejected: pd.DataFrame,
    output_path: Path,
) -> None:
    """导出 Excel 报表，包含两个工作表：订单明细、异常记录。"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    detail_cols = ["订单号", "客户", "产品", "数量", "单价", "金额", "日期", "source_file", "source_row"]
    reject_cols = ["订单号", "客户", "产品", "数量", "单价", "日期", "source_file", "source_row", "error"]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        valid[detail_cols].to_excel(writer, sheet_name="订单明细", index=False)
        rejected[reject_cols].to_excel(writer, sheet_name="异常记录", index=False)


# ---------------------------------------------------------------------------
# 总流程
# ---------------------------------------------------------------------------
def process_directory(input_dir: Path, output_path: Path) -> dict:
    """串联全部步骤，返回统计信息字典。"""
    raw = read_order_files(input_dir)
    cleaned = clean_orders(raw)
    valid, invalid = validate_orders(cleaned)
    kept, duplicates = deduplicate_orders(valid)
    rejected = pd.concat([invalid, duplicates], ignore_index=True)

    export_report(kept, rejected, output_path)

    return {
        "读取行数": len(raw),
        "有效行数": len(kept),
        "校验失败": len(invalid),
        "重复订单": len(duplicates),
        "输出文件": str(output_path),
    }


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print("用法: python -m app.excel_batch <输入目录> <输出报表.xlsx>")
        return 1
    stats = process_directory(Path(argv[1]), Path(argv[2]))
    for key, value in stats.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
