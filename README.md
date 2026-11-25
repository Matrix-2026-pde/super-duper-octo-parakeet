# 企业 Excel 订单数据管理系统

最小化的订单数据闭环项目，用于学习基础全栈开发与业务自动化。

当前进度：**阶段一（Python Excel 批量处理）**

## 快速开始

```powershell
python -m pip install -r requirements.txt
python scripts/make_sample_data.py
python -m app.excel_batch data/samples data/output/report.xlsx
python -m pytest
```

详见 `AGENTS.md`。
