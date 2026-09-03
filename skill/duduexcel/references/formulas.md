# 公式白名单与踩坑

## 可安全使用

`SUM` `AVERAGE` `SUMIFS` `COUNTIFS` `COUNT` `COUNTA` `INDEX` `MATCH`
`IF` `IFERROR` `SUMPRODUCT` `RANK` `VLOOKUP` `HLOOKUP` `MIN` `MAX`
`ROUND` `ABS` `PMT` `NPV` `IRR`

## 必须加 `_xlfn.` 前缀（否则 #NAME?）

| 函数 | 写法 |
|---|---|
| TEXTJOIN | `_xlfn.TEXTJOIN` |
| CONCAT | `_xlfn.CONCAT` |
| IFS | `_xlfn.IFS` |
| SWITCH | `_xlfn.SWITCH` |
| MAXIFS | `_xlfn.MAXIFS` |
| MINIFS | `_xlfn.MINIFS` |

原因：openpyxl 把公式逐字写进 XML，而 Excel 内部存储带此前缀（UI 里隐藏）。

## 禁用：溢出数组函数

`XLOOKUP` `XMATCH` `SORT` `FILTER` `UNIQUE` `SEQUENCE`

原因：它们依赖 spill 元数据，openpyxl 写出的文件没有该元数据，
结果只有左上角单元格有值，**且错误扫描报 0 错误（静默截断）**。

替代方案：
| 需求 | 替代 |
|---|---|
| 查找 | `INDEX` + `MATCH` |
| 排序 | 在 Python 里排好再写死 |
| 筛选 | 在 Python 里筛完再写死 |
| 去重 | 在 Python 里去完再写死 |

## 诊断技巧

- LibreOffice 解析不了的公式会被**回写成小写** —— 这是 `#NAME?` 的快速信号。
- 含空格的表名跨表引用必须加引号：`='销售 明细'!B2`（不加 → `#VALUE!`）。

## 七类错误速查

| 错误 | 常见原因 |
|---|---|
| `#VALUE!` | 类型不匹配（文本参与算术） |
| `#DIV/0!` | 分母为零 |
| `#REF!` | 引用区域被删除 |
| `#NAME?` | 函数名不存在 / 缺 `_xlfn.` 前缀 / 表名未加引号 |
| `#NULL!` | 区域运算符用错（空格代替逗号） |
| `#NUM!` | 数值越界（如负数的平方根） |
| `#N/A` | 查找失败 |

## 自检顺序

1. `write_cells` 写公式
2. `scan_formula_errors` 体检（快，不需要 LibreOffice）
3. 装了 LibreOffice 再跑 `recalculate` 重算
4. 抽样 2–3 个单元格，人工核对数值是否符合预期
5. 确认无误再铺开整片区域
