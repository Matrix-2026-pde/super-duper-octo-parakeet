"""生成演示用的脏数据 Excel，用于验证阶段一的清洗、校验、去重。

用法（项目根目录）：
    python scripts/make_sample_data.py
输出：
    data/samples/orders_2026_01.xlsx
    data/samples/orders_2026_02.xlsx
"""

from pathlib import Path

import pandas as pd

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "data" / "samples"

COLUMNS = ["订单号", "客户", "产品", "数量", "单价", "日期"]


def build_january() -> pd.DataFrame:
    """1 月文件：包含正常数据、带空白的文本、非法数量、缺失日期。"""
    return pd.DataFrame(
        [
            ["SO-1001", " 华北贸易 ", "显示器", 2, 999.0, "2026-01-05"],
            ["SO-1002", "东方科技", "键盘", 10, 129.5, "2026-01-07"],
            ["SO-1003", "蓝海公司", "鼠标", 0, 59.0, "2026-01-08"],       # 数量为 0 -> 非法
            ["SO-1004", "星辰集团", "路由器", 3, 299.0, None],            # 日期缺失 -> 非法
            ["SO-1005", "   ", "交换机", 1, 450.0, "2026-01-10"],          # 客户为空白 -> 非法
            ["SO-1001", "华北贸易", "显示器", 2, 999.0, "2026-01-05"],    # 与第 2 行重复
        ],
        columns=COLUMNS,
    )


def build_february() -> pd.DataFrame:
    """2 月文件：包含正常数据、单价为文本、订单号跨文件重复。"""
    return pd.DataFrame(
        [
            ["SO-2001", "远航物流", "显示器", 5, 980.0, "2026-02-03"],
            ["SO-2002", "金桥电子", "硬盘", 4, "468.00", "2026-02-09"],   # 单价为文本，清洗后可用
            ["SO-2003", "晨光文具", "打印机", -1, 1200.0, "2026-02-12"],  # 数量为负 -> 非法
            ["SO-1003", "蓝海公司", "鼠标", 20, 59.0, "2026-02-15"],      # 与 1 月订单号重复
            ["SO-2004", "北辰商贸", "耳机", 6, 199.0, "not-a-date"],     # 日期无法解析 -> 非法
        ],
        columns=COLUMNS,
    )


def main() -> None:
    SAMPLE_DIR.mkdir(parents=True, exist_ok=True)
    build_january().to_excel(SAMPLE_DIR / "orders_2026_01.xlsx", index=False)
    build_february().to_excel(SAMPLE_DIR / "orders_2026_02.xlsx", index=False)
    print(f"已生成示例数据到: {SAMPLE_DIR}")


if __name__ == "__main__":
    main()
