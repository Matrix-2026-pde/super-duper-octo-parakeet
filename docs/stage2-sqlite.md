# 阶段二讲解：SQLite 数据库与基础 SQL

本文档对应代码 `app/db.py`、`app/excel_batch.py` 中的 `run_pipeline`，以及测试 `tests/test_db.py`。目的是让每一行代码都能讲清楚原理、意义和取舍。

---

## 一、为什么需要数据库

阶段一的结果是一个 Excel 报表。它适合“一次性导出”，但有三个问题：

| 问题 | Excel 报表 | 数据库 |
|---|---|---|
| 修改单条订单 | 要重开文件、找行、改、再保存 | `UPDATE ... WHERE order_id = ?` 一句完成 |
| 重复导入同一批数据 | 容易产生重复行 | 主键约束自动拒绝或更新重复 |
| 多条件查询、汇总 | 需要手工筛选或公式 | `WHERE` / `GROUP BY` 直接计算 |

所以阶段二的目标是：把“清洗后的有效数据”沉淀到一个可增删改查的存储中。

---

## 二、SQLite 的核心原理

SQLite 是一个**嵌入式**数据库：没有独立的服务进程，整个数据库就是磁盘上的一个 `.db` 文件。Python 标准库自带 `sqlite3` 模块，因此**无需额外安装依赖**。

几个关键概念：

1. **表（table）**：行和列的集合。本项目只有一张表 `orders`。
2. **主键（PRIMARY KEY）**：每行的唯一标识。这里用 `order_id`，数据库保证它不重复。
3. **约束（CONSTRAINT）**：数据库层面的规则，例如 `NOT NULL`、`CHECK (quantity > 0)`。即使程序有 bug，数据库也会拒绝非法数据。
4. **索引（INDEX）**：为查询加速的“目录”。这里给 `customer` 和 `order_date` 建索引，按客户或日期查询时不必逐行扫描。
5. **事务（TRANSACTION）**：一组操作“要么全部成功，要么全部撤销”。这是数据一致性的基础。

---

## 三、建表语句 `SCHEMA`

```sql
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
```

逐项说明：

- `IF NOT EXISTS`：表已存在时不报错。使脚本可以重复运行（幂等）。
- `order_id TEXT PRIMARY KEY`：订单号作为主键。阶段一已经按订单号去重，这里再次由数据库兜底。
- `CHECK (quantity > 0)`：数据库级的“数量必须为正”。与阶段一的 `validate_orders` 形成双重保护。
- `REAL` 与 `TEXT`：SQLite 是“动态类型”，但声明类型会影响存储与比较方式。金额用 `REAL`（浮点），日期用 `TEXT`（见下节）。
- `source_file` / `source_row`：保留来源信息，便于回溯到原 Excel 的哪个文件、哪一行，与阶段一的约定一致。

**为什么日期存成 TEXT？** SQLite 没有独立的日期类型。存成 `YYYY-MM-DD` 文本后，字符串的字典序就等于时间顺序，因此 `ORDER BY order_date` 与 `BETWEEN` 都能正确工作，而且可读性好。

---

## 四、UPSERT：插入或更新

```sql
INSERT INTO orders (...) VALUES (?, ?, ...)
ON CONFLICT (order_id) DO UPDATE SET
    customer = excluded.customer, ...
```

- `INSERT` 的 `VALUES` 中用 `?` 占位符，值由 `executemany` 传入。
- `ON CONFLICT (order_id)`：当主键已存在时，不报错，而是转为 `DO UPDATE`。这就是 **UPSERT**（Update + Insert）。
- `excluded` 是特殊表名，指“本次尝试插入但因冲突被拒的那一行”。

**意义**：同一批 Excel 重复导入，或新文件里出现已有订单的更新版本，都能安全处理，不会因主键冲突而中断。

---

## 五、参数化查询（防 SQL 注入）

错误写法：

```python
conn.execute(f"SELECT * FROM orders WHERE order_id = '{order_id}'")  # 危险
```

如果 `order_id` 来自用户输入或 Excel，恶意内容（例如 `' OR '1'='1`）会改变 SQL 语义，这就是 **SQL 注入**。

正确写法：

```python
conn.execute("SELECT * FROM orders WHERE order_id = ?", (order_id,))
```

`?` 是占位符，数据与 SQL 分开传输，数据库只把它当作“值”，永远不当作“代码”。**本项目所有含外部数据的语句都使用这种写法。**

---

## 六、事务与原子性

```python
with conn:
    conn.executemany(INSERT_SQL, records)
```

`sqlite3` 中 `with conn:` 的语义：
- 块内全部成功 → 自动 `COMMIT`，数据落盘；
- 块内抛出异常 → 自动 `ROLLBACK`，所有改动撤销。

测试 `test_import_is_atomic_on_failure` 验证了这一点：第二行数量为 0 违反 `CHECK`，整批回滚，数据库中一条都不留。这避免了“导入一半”的脏状态。

---

## 七、类型转换：为什么要 `to_records`

pandas 的值类型（`numpy.int64`、`pandas.Timestamp`）**sqlite3 不认识**，直接传入会报错。因此 `to_records` 把每个值转换为原生类型：

| pandas 类型 | 转换后 | 原因 |
|---|---|---|
| `numpy.int64` | `int` | sqlite3 只接受 Python 原生整数 |
| `pandas.Timestamp` | `'YYYY-MM-DD'` 字符串 | SQLite 没有日期类型，统一为可排序文本 |
| 数值/文本 | `float` / `str` | 与表结构类型一致 |

---

## 八、基础 SQL 详解

### 1. 按主键查询
```sql
SELECT * FROM orders WHERE order_id = ?;
```
主键查询最快，数据库直接定位到唯一一行。

### 2. 条件查询 + 排序
```sql
SELECT * FROM orders WHERE customer = ? ORDER BY order_date, order_id;
```
`WHERE` 过滤行，`ORDER BY` 排序。使用了 `customer` 索引加速过滤。

### 3. 分组汇总
```sql
SELECT customer,
       COUNT(*)              AS order_count,
       SUM(quantity)         AS total_quantity,
       ROUND(SUM(amount), 2) AS total_amount
FROM orders
GROUP BY customer
ORDER BY total_amount DESC;
```
- `GROUP BY customer`：把同一客户的行归为一组。
- `COUNT / SUM`：对每组做聚合。
- `ROUND(..., 2)`：金额保留两位小数，便于展示。

这正是 Excel 数据透视表的 SQL 版本。

### 4. 修改（UPDATE）
```sql
UPDATE orders SET quantity = ?, amount = ? * unit_price WHERE order_id = ?;
```
注意：金额是**派生字段**（= 数量 × 单价），修改数量时必须同时更新金额，否则报表会自相矛盾。`cur.rowcount` 告诉我们是否真的命中了一行。

### 5. 删除（DELETE）
```sql
DELETE FROM orders WHERE order_id = ?;
```
同样用 `rowcount` 判断是否删除成功。**注意**：没有 `WHERE` 的 `DELETE` 会清空整张表，本项目每个删除都带主键条件。

---

## 九、Python 与 SQL 的分工

| 层 | 职责 | 本项目对应 |
|---|---|---|
| 数据库约束 | 最后一道防线，拒绝非法数据 | `CHECK`、`PRIMARY KEY`、`NOT NULL` |
| Python 校验 | 友好的错误信息、可统计 | `validate_orders`（阶段一） |
| SQL 语句 | 存储、查询、汇总 | `app/db.py` 中的各函数 |

理念：**业务规则放在哪里都要有，且越靠近数据越可靠。**

---

## 十、运行与验收

```powershell
python -m pip install -r requirements.txt
python scripts/make_sample_data.py --rows 2000
python -m app.db data/samples data/output/orders.db
python -m pytest
```

预期：
- 导入命令输出：读取 6000 行，有效 3944，写入 3944，数据库总数 3944；
- pytest：23 项全部通过。

可以用 SQLite 命令行或 DB Browser for SQLite 打开 `data/output/orders.db` 自行验证：

```sql
SELECT customer, COUNT(*) FROM orders GROUP BY customer ORDER BY 2 DESC LIMIT 5;
```

---

## 十一、常见问题

- **为什么不用 ORM（如 SQLAlchemy）？** 阶段二的目标是掌握 SQL 本身。ORM 会隐藏 SQL，初学阶段不利于理解原理。
- **为什么金额存 REAL 而不是 DECIMAL？** SQLite 没有定点小数类型。金额要求精确时可存“分”为整数，但本项目规模下 `ROUND(..., 2)` 已足够，并在文档中说明这一取舍。
- **并发写入怎么办？** SQLite 同一时刻只允许一个写入者，适合单机、低并发场景，这也是本项目选择它的原因。
