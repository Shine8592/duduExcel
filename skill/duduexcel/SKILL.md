---
name: duduexcel
description: Excel 表格处理（.xlsx/.xlsm）。当用户需要读取、分析、批量写入、格式化、绘制图表或重算公式时使用——包括"看看这个表有什么""统计一下销售额""筛选某部门的数据""做个柱状图""套一下人民币格式"。用中文或英文描述均可。Do NOT trigger when 主要交付物是 Word 文档、PPT、PDF、独立 Python 脚本、数据库管道或 Google Sheets API 集成。
license: MIT
compatibility: 需要 Python 3.10+、openpyxl、pandas；公式重算（recalculate）还需本机安装 LibreOffice
metadata:
  version: "0.2.0"
  transport: stdio
---

# duduExcel 使用指南

## 工具路由（按任务选工具，别选错）

| 任务 | 用哪个工具 |
|---|---|
| 先看表结构 | `workbook_info` → `sheet_profile` |
| 取具体数据 | `read_range`（分页，别一次全读） |
| 问"有多少条/几条满足…" | `filter_count` |
| 求和、均值、分组统计 | `aggregate` |
| 排行榜 Top N | `top_n` |
| 写多个单元格 | `write_cells`（**一次传完，禁止循环单格写**） |
| 数字格式（¥、%） | `set_number_format` |
| 中文表格美化 | `apply_chinese_style` |
| 图表 | `add_chart` |
| 公式重算与体检 | `recalculate` / `scan_formula_errors` |
| 写错了想撤销 | `revert_last_write` |

## 铁律（违反会出事故）

1. **绝不把整表读进上下文**。默认 `limit=200`；`truncated=true` 时按 `hint` 用 `offset` 分页，不要盲目调大 limit。
2. **批量写入，不要循环**。一次 `write_cells` 传完所有单元格。
3. **用公式，不要写死计算结果**。写 `=SUM(C2:C7)`，不要自己算完写数字。
4. **交付前必须验算**。写完公式后用 `scan_formula_errors` 体检；有 LibreOffice 时用 `recalculate` 重算。
   - 重算通过 ≠ 结果正确：错一位的引用会产出"干净但错"的文件。**先写 2–3 个公式验证取值，再铺开整片区域**。
5. **写操作前自动备份**，`revert_last_write` 可回滚最近一次。

## 公式写法约束

- **能用的常规函数**：`SUM` `AVERAGE` `SUMIFS` `COUNTIFS` `INDEX` `MATCH` `IFERROR` `SUMPRODUCT` `RANK` `VLOOKUP`
- **6 个较新函数必须加 `_xlfn.` 前缀**，否则显示 `#NAME?`：
  `_xlfn.TEXTJOIN` `_xlfn.CONCAT` `_xlfn.IFS` `_xlfn.SWITCH` `_xlfn.MAXIFS` `_xlfn.MINIFS`
- **禁用溢出数组函数**：`XLOOKUP` `XMATCH` `SORT` `FILTER` `UNIQUE` `SEQUENCE`
  → 替代：查找用 `INDEX`/`MATCH`；排序筛选去重**在 Python 里算完再写死单元格**。
- 含空格的表名跨表引用要加引号：`='销售 明细'!B2`

## openpyxl 常见坑（本服务已规避，但你直接用库时要注意）

1. 读公式与读值**要加载两次**（`data_only=True` 得值、默认得公式）。
2. `data_only=True` 后保存会**把公式永久变成字面量**（不可逆）。
3. 刚写入未重算的公式，读回来是 `None` —— 这是预期，不是数据丢失。
4. 合并单元格只有左上角可写，其余只读。
5. `.xlsm` 要保留 VBA。
6. 迭代行得到的是 Cell 对象，要取 `.value`。

## 中文场景规范

- 字体：默认**微软雅黑**（屏幕阅读）或宋体（正式/打印），一般不用 Arial。
- 货币：`¥#,##0`；零显示为短横用 `¥#,##0;(¥#,##0);-`。
- 百分比：`0.0%`，**值必须存小数**（0.15 显示 15.0%；存 15 会显示 1500.0%）。
- 估值倍数：`0.0x`。年份当文本存（`"2024"` 而非 `2,024`）。

## 财务建模配色（行业通行惯例）

| 颜色 | 含义 |
|---|---|
| 蓝字 | 硬编码输入 / 情景开关 |
| 黑字 | 公式 |
| 绿字 | 跨工作表引用 |
| 红字 | 跨文件引用 |
| 黄底 | 关键假设 / 需用户填写 |

## 深入参考

- 样式与配色细则 → `references/style.md`
- 公式白名单与踩坑 → `references/formulas.md`
- 图表做法与保真度 → `references/charts.md`

## 已知限制

- 未装 LibreOffice 时 `recalculate` 会明确降级（不会静默假装成功）；此时公式无缓存值，读回 `None`。
- 透视表、条件格式尚未支持（图表已支持）。
- 单次写入上限 10 万单元格。
