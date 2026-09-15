# 示例

本目录包含一个用于快速体验 duduExcel 的演示文件。

## 快速开始（30 秒）

```bash
pip install "duduexcel[analysis]"
python examples/make_demo.py      # 重新生成 demo.xlsx（已内置，可跳过）
```

把它接到你的 MCP 客户端后，试试这些指令：

```
先看清这张表：examples/demo.xlsx 的「需求清单」
```
→ 调用 `workbook_info` + `sheet_profile`，**不把任何数据行读进上下文**

```
读一下需求清单，哪些条目已经取消了？
```
→ 调用 `read_range`，返回带格式语义的视图：
```
A3: 旧版导出 [S] | B3: 李四 [S] | C3: 已取消 [S]     ← [S] = 删除线 = 已取消
A4: 司机点名      | B4: 王五      | C4: 待审阅 [HL:FFFF00]  ← 黄底 = 待审阅
```
> **其他 Excel MCP 会把"C3=已取消"和"C2=已批准"读成一样的字符串**，看不到作者用格式留下的意图。

```
按区域统计销售额
```
→ 调用 `aggregate(group_by=["区域"])`，服务端算完只回传 2 行结果，而非 6 行原始数据

```
给「销售」表做一个可交互透视表
```
→ 调用 `create_interactive_pivot`，生成**真正的 Excel PivotTable**（可拖拽字段、刷新），
  不是静态汇总表

```
合计那行公式算一下，看看结果对不对
```
→ 调用 `recalculate`（需装 LibreOffice），把 `=SUM(C2:C7)` 算出真实值

---

## demo.xlsx 里埋了什么

| 位置 | 埋点 | 用途 |
|---|---|---|
| 「需求清单」A3:C3 | **删除线** | 演示格式语义（已取消）|
| 「需求清单」C4 | **黄底** | 演示格式语义（待审阅）|
| 「需求清单」第 5 行 | **隐藏行** | 演示隐藏内容处理（默认跳过并报告）|
| 「销售」表 | 中文表头 + 数值 | 演示分析/透视/图表 |
| 「销售」C8:D8 | `=SUM(...)` 公式 | 演示公式重算（写入时无缓存值）|

## 自己重造演示文件

```bash
python examples/make_demo.py
```

脚本见 [`make_demo.py`](make_demo.py)，全部用 openpyxl 构造，可自由修改。

---

## 注意

- `demo.xlsx` 里的「透视表演示」工作表是通过 `create_interactive_pivot` 现场生成的，
  重新运行 `make_demo.py` 会**重置**它（因为脚本会覆盖整个文件）。
- 交互透视表有个已知限制：**生成后若再用 openpyxl 保存该文件，透视表会丢失**。
  所以生成透视表请放在最后一步。
