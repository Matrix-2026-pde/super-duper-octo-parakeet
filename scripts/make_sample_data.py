"""生成演示用的批量脏数据 Excel，用于验证阶段一的清洗、校验、去重。

用法（项目根目录）：
    python scripts/make_sample_data.py              # 默认每个文件 300 行，共 3 个文件
    python scripts/make_sample_data.py --rows 2000  # 指定每个文件行数

设计要点：
- 使用固定随机种子（SEED），同样参数每次生成的数据完全一致，便于测试复现。
- 按固定比例注入脏数据（空客户、非法数量、缺失日期、文本单价、重复订单号等），
  比例由 DIRTY_RATE 控制，用来模拟真实业务中的数据质量问题。
- 表头写法在不同文件间故意不同（如 "订单编号" / "order_id"），以验证别名映射。
"""

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "data" / "samples"

SEED = 20260101          # 固定随机种子，保证可复现
DIRTY_RATE = 0.08        # 每种脏数据的注入概率（约 8%）

CUSTOMERS = ["华北贸易", "东方科技", "蓝海公司", "星辰集团", "远航物流",
             "金桥电子", "晨光文具", "北辰商贸", "云帆实业", "松果科技"]
PRODUCTS = ["显示器", "键盘", "鼠标", "路由器", "交换机", "硬盘", "打印机", "耳机", "摄像头", "音箱"]

# 不同文件使用不同表头写法，用来验证 COLUMN_ALIASES 映射
HEADER_STYLES = [
    ["订单号", "客户", "产品", "数量", "单价", "日期"],
    ["订单编号", "客户名称", "商品", "件数", "单价(元)", "下单日期"],
    ["order_id", "customer", "product", "quantity", "unit_price", "date"],
]

COLUMNS = HEADER_STYLES[0]


def build_file(rng: np.random.Generator, rows: int, file_index: int) -> pd.DataFrame:
    """生成一个文件的数据：先造正常行，再按比例注入脏数据。"""
    prefix = f"SO-{file_index + 1}"
    records = []
    for i in range(rows):
        records.append([
            f"{prefix}-{i:05d}",
            rng.choice(CUSTOMERS),
            rng.choice(PRODUCTS),
            int(rng.integers(1, 50)),
            round(float(rng.uniform(10, 2000)), 2),
            (pd.Timestamp("2026-01-01") + pd.Timedelta(days=int(rng.integers(0, 330)))).strftime("%Y-%m-%d"),
        ])

    # 先把所有列设为 object，允许后续注入字符串/None 等脏值，避免 pandas dtype 警告
    df = pd.DataFrame(records, columns=COLUMNS, dtype=object)

    # 按 DIRTY_RATE 注入脏数据（每种类型独立抽样）
    _inject(rng, df, "客户", lambda: "   ")                 # 空白客户 -> 校验失败
    _inject(rng, df, "数量", lambda: int(rng.choice([0, -3])))  # 非正数量 -> 校验失败
    _inject(rng, df, "日期", lambda: None)                   # 日期缺失 -> 校验失败
    _inject(rng, df, "日期", lambda: "not-a-date")           # 日期无法解析 -> 校验失败
    _inject(rng, df, "单价", lambda: f"¥{float(rng.uniform(10, 2000)):,.2f}")  # 文本单价 -> 清洗后可用
    # 重复订单号：把某些行的订单号改成文件中第一行的订单号 -> 去重
    first_id = df.loc[0, "订单号"]
    _inject(rng, df, "订单号", lambda: first_id)
    return df


def _inject(rng, df, column, value_factory):
    """按 DIRTY_RATE 随机选中行，写入脏值。"""
    mask = rng.random(len(df)) < DIRTY_RATE
    for idx in np.flatnonzero(mask):
        df.loc[idx, column] = value_factory()


def main() -> None:
    parser = argparse.ArgumentParser(description="生成批量演示订单数据")
    parser.add_argument("--rows", type=int, default=300, help="每个文件的行数，默认 300")
    args = parser.parse_args()

    SAMPLE_DIR.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(SEED)

    for i, header in enumerate(HEADER_STYLES):
        df = build_file(rng, args.rows, i)
        df.columns = header  # 按该文件的表头写法输出
        path = SAMPLE_DIR / f"orders_batch_{i + 1:02d}.xlsx"
        df.to_excel(path, index=False)
        print(f"已生成 {path.name}：{len(df)} 行")

    print(f"示例数据目录: {SAMPLE_DIR}")


if __name__ == "__main__":
    main()
