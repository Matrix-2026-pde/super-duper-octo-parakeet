# AGENTS.md

## 项目目标

一个最小的企业 Excel 订单数据管理系统，用于学习基础全栈开发与业务自动化的完整闭环：
批量读取 Excel 订单 → 清洗、校验、去重 → 写入 SQLite → 网页查询/修改 → 导出 Excel 报表。

目标是掌握流程，不追求功能复杂度。

## 技术栈（固定，不得更换或追加）

- 原生 HTML / CSS / JavaScript
- Flask
- SQLite
- Pandas
- openpyxl
- pytest
- Git

不复用现有完整项目，不引入其他框架，不增加未要求的功能。

## 开发阶段

1. Python Excel 批量处理（当前阶段）
2. SQLite 数据库及基础 SQL
3. Flask 后端接口
4. 原生 Web 前端
5. 测试与工程交付

**流程规则**：每个阶段必须说明核心原理、列出涉及文件、给出运行命令与验收方法。
完成当前阶段后**立即停止**，等待用户运行、阅读代码并确认，才能进入下一阶段。

## 目录结构

```
企业自动化数据分析平台/
├── AGENTS.md              # 项目说明与协作规则
├── README.md              # 对外说明（工程交付阶段完善）
├── requirements.txt       # Python 依赖
├── pytest.ini             # pytest 配置
├── app/
│   ├── __init__.py
│   └── excel_batch.py     # 阶段一：批量读取、清洗、校验、去重、导出
├── scripts/
│   └── make_sample_data.py  # 生成演示用脏数据 Excel
├── tests/
│   └── test_excel_batch.py  # 阶段一单元/集成测试
└── data/
    ├── samples/           # 演示输入 Excel（可提交）
    └── output/            # 处理结果（忽略，不提交）
```

## 常用命令（在项目根目录执行）

```powershell
python -m pip install -r requirements.txt
python scripts/make_sample_data.py
python -m app.excel_batch data/samples data/output/report.xlsx
python -m pytest
```

## 编码约定

- 注释、docstring、界面文字使用中文。
- 清洗、校验、去重逻辑写成纯函数，输入 DataFrame，输出 DataFrame，便于 pytest。
- 每条错误记录必须保留 `source_file` 与 `source_row`，便于回溯到原 Excel 行。
- 业务数据文件不提交到 Git（`data/output/` 已忽略）。

## 当前状态

- 阶段一：已完成代码（含表头别名映射、数字清洗、批量样例数据），待用户运行与确认。
- 阶段二及以后：未开始。

## 阶段一已知约定

- 表头别名见 `app/excel_batch.py` 的 `COLUMN_ALIASES`，新增写法只改此处。
- 样例数据由 `scripts/make_sample_data.py` 固定种子生成，默认每文件 300 行、共 3 个文件。
- 数据是否合格以 `validate_orders` 为准，不在清洗阶段静默丢弃。
