# super-duper-octo-parakeet

企业 Excel 订单数据管理系统：批量读取 Excel 订单 → 清洗、校验、去重 → 写入 SQLite → 网页查询/修改 → 导出 Excel 报表。

当前进度：**阶段一（Python Excel 批量处理）**

## 快速开始

```powershell
python -m pip install -r requirements.txt
python scripts/make_sample_data.py
python -m app.excel_batch data/samples data/output/report.xlsx
python -m pytest
```

详见 `AGENTS.md`。
