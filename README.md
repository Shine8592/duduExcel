# 📊 duduExcel

面向 AI Agent 的 Excel MCP 服务 —— **上下文高效 + 安全读写 + 中文场景 + 公式验证**。

```
Agent ──MCP(stdio)──► duduExcel ──► openpyxl/pandas ──► .xlsx
```

## 为什么又一个 Excel MCP？

调研了 GitHub 上 10+ 个同类项目后，发现它们分裂成两个对立阵营，没人同时做到"能写"和"不炸上下文"：

| 阵营 | 代表 | 能写 | 上下文友好 | 问题 |
|---|---|---|---|---|
| 读写全能型 | haris-musa、knorq | ✅ | ❌ | 整表读进上下文就爆了 |
| 原子分析型 | jwadow、jdatamunch | ❌ 只读 | ✅ | 改不了文件 |
| 方法论型 | Anthropic 官方 xlsx Skill | ✅（靠写代码） | ✅ | 靠 Agent 自己写 Python，不稳定 |

**duduExcel：把三者优点合到一个项目** —— 服务端原子分析（省 token）+ 完整读写 + 内置方法论 Skill。

## ✨ 能力（17 个工具）

| 层 | 工具 | 说明 |
|---|---|---|
| **探查** | `workbook_info` | 表清单、行列数、文件大小 |
| | `sheet_profile` | 列画像：类型/空值率/唯一数/Top值/统计（一次替代十几次调用） |
| **分析** | `filter_count` | 条件计数（14 种运算符），只回传数字 |
| | `aggregate` | sum/mean/count/… 支持分组与过滤 |
| | `top_n` | 排行榜，只回传前 N 行 |
| **读写** | `read_range` | 分页读取，默认 limit=200 防上下文溢出 |
| | `write_cells` | **批量**写入，一次调用完成；`=` 开头即公式 |
| **验证** ⭐ | `recalculate` | 公式重算 + 外链熔断（差异化，竞品多不支持） |
| | `scan_formula_errors` | 扫描 7 类公式错误 |
| **中文** ⭐ | `apply_chinese_style` | 微软雅黑表头、中文列宽自适应、冻结首行 |
| | `set_number_format` | `¥#,##0` / `0.0%` / `0.0x` 等内置格式 |
| **图表** ⭐ | `add_chart` | bar/line/pie/scatter（官方 skill 无图表指导，knorq 明确不支持） |
| **高级** | `create_pivot` | 透视汇总表（静态，诚实标注不可交互） |
| | `add_conditional_format` | 数据条/色阶/阈值高亮/重复值（openpyxl 原生规则） |
| | `compare_sheets` | 两表按关键列比对，只回差异摘要 |
| | `join_sheets` | 两表关联（left/right/inner/outer），只回前 N 行 |
| **安全** | `revert_last_write` | 回滚最近一次写入 |

⭐ = 差异化能力

## 🔑 汲取的设计（附来源）

| 设计点 | 来源 | 落地 |
|---|---|---|
| 一次调用替代 N 次试探 | jwadow `get_data_profile` | `sheet_profile` |
| 服务端原子操作（results, not rows） | jwadow | `filter_count`/`aggregate`/`top_n` |
| 分页 + 截断诚实报告 | jwadow、官方 `recalc.py` | `read_range` 返回 `truncated`+`hint` |
| `_meta.tokens_saved` 自报节省 | jdatamunch | 每个分析工具的 `_meta` |
| 结果附 Excel 公式（可复现） | jwadow | `filter_count`/`aggregate`/`top_n` |
| TSV 输出（便于粘回 Excel） | jwadow | `aggregate`/`top_n` |
| 批量接口，禁止循环调用 | knorq | `write_cells` |
| 路径白名单 + 拒目录穿越 | haris-musa | `DUDU_EXCEL_ROOT` |
| 外链熔断（不可逆保护） | 官方 `recalc.py` | `recalculate` 的 force 机制 |
| 文件指纹防静默失败 | 官方 `recalc.py` | `recalculate` |
| 诚实的 Limitations | knorq | 见下 |

## 🚀 安装与挂载

```bash
cd duduExcel
pip install -e ".[analysis]"     # analysis 可选（提供 pandas）
```

opencode（`~/.config/opencode/opencode.jsonc`）：

```jsonc
"duduexcel": {
  "type": "local",
  "command": ["<你的python路径>", "-u", "-m", "duduexcel"],
  "enabled": true,
  "cwd": "E:\\工作类\\研发\\duduExcel"
}
```

其他 MCP 客户端（Claude Code / Cursor / Cline）同理，指向 `python -m duduexcel`。
**配置改完需重启客户端生效。**

## 🔒 安全

- **路径白名单**：设 `DUDU_EXCEL_ROOT` 后，所有路径必须在根目录内，拒绝绝对路径与 `..` 穿越
- **写前自动备份**：任何写入前生成 `.xlsx.bak`，备份失败则中止写入
- **失败自动回滚**：写入异常时自动还原，不留半成品
- **外链熔断**：检测到外链缓存丢失时拒绝重算（否则会永久破坏外链），需显式 `force=true`
- **本地优先**：stdio 传输，文件不出本机

## 🔄 汲取的新设计（第二轮调研）

调研 `VOYAGER-Inc/excel-vision-mcp` 后发现一个 openpyxl 生态的普遍盲区，
而 duduExcel 原本也有：**内嵌图片、格式语义、隐藏行列全部丢失**。
作者用删除线表示"已取消"、黄底表示"待审阅"，而纯文本读取会把它们抹平——
模型看到的是"字符串表格"，不是"作者想表达的意思"。

已补齐：

| 汲取点 | 来源 | 落地 |
|---|---|---|
| **格式语义标记** | excel-vision-mcp | `read_range` 返回 `[B]/[I]/[S]/[HL:色]/[C:色]/[M]`；**只在真用了格式时才附加**，朴素表零额外 token |
| **隐藏内容智能处理** | excel-vision-mcp | 默认跳过隐藏行列（作者隐藏=不想让你看），但**被可见公式引用的隐藏单元格仍保留**并标 `[HIDDEN-REF]`；始终报告跳过数量，绝不静默丢弃 |
| **内嵌图片不丢失** | excel-vision-mcp | 新增 `list_images`（零依赖扫描 `xl/media/`），`workbook_info` 也报告图片数 |
| **原子保存** | excel-vision-mcp | 写临时文件成功后再替换目标，失败的写入永不损坏原文件 |
| **多目录沙箱** | excel-vision-mcp | `DUDU_EXCEL_ROOT` 支持 `;` 分隔多个目录 |

## ⚠️ 已知限制（诚实清单）

- **公式重算需要 LibreOffice**：未安装时 `recalculate` 明确降级并给出安装指引（不静默假装成功）。
  此时写入的公式无缓存值，`read_range` 读回 `None` —— 这是预期行为，非数据丢失。

  ✅ **已在本机打通验证**（LibreOffice 26.8）：`recalculate` 实测 **4 秒**完成重算，
  `=SUM(销售!C2:C7)` 正确算出 **575** 并落盘（回归测试见 `tests/test_m8_recalc.py`）。

  ⚠️ **中文路径坑（已修复）**：LibreOffice 在中文路径下**原地覆盖**文件会失败
  （`SfxBaseModel::impl_store failed: 0x4c0c`）。本实现改为先输出到纯 ASCII 临时目录再移回，
  因此在 `E:\工作类\研发\` 这类中文路径下也能正常重算。
- **不支持可交互透视表**：`create_pivot` 生成的是静态汇总表（数值正确且已验证，如华东 365 / 华南 210），
  但 openpyxl 无法创建真正的 PivotTable 对象（实测 `ws._pivots` 为空），因此不可点击交互。
- ~~条件格式只能写入不能读取~~ —— **已更正**：实测条件格式**可以完整读回**。
  `list_conditional_formats` 能返回作用区域/类型/运算符/阈值/填充色/优先级（实测 4 条规则全部读回）。
  此前我照抄了竞品 knorq 的 Known Limitations 而未亲自验证，这是错误的。
- **`.xlsm` 宏**：读取保留 VBA，写入不保证
- 单次写入上限 **10 万单元格**
- 小表上 `_meta.tokens_saved` 节省不明显属正常（省 token 的收益随表增大而放大）

## 🧪 测试

```bash
python tests/test_smoke.py       # M1 读写与分页
python tests/test_m2.py          # M2 服务端分析
python tests/test_m34.py         # M3 重算降级 + M4 中文样式与图表
python tests/test_m6.py          # M6 透视表/条件格式/多表关联
python tests/test_edge_cases.py  # 边界回归（防 BUG 复发）
python tests/test_m7.py          # M7 格式语义/隐藏处理/图片/原子保存
```

6 个测试文件 **全部通过**，另有 **20 项 MCP stdio 端到端验证**（覆盖全部 19 个工具）。

## 🧠 踩过的坑（值得记住）

1. **openpyxl `read_only=True` 模式不加载行列维度** —— `ws.row_dimensions` 为空，
   隐藏检测会**静默失效**。改为直接解析 sheet XML（零内存且准确）。
2. **mcp 2.x 的 schema 生成器会丢弃 `str | None`（PEP 604）标注的参数** ——
   表现为"工具传了参数却不生效"。统一改用 `Optional[str]` 才正常。
3. **图表 `Reference(range_string="B2:B5")` 要求 `表名!A1:B2` 形式** ——
   中文表名还需引号包裹。改用 `range_boundaries()` 解析成行列参数。
4. **重复 `@mcp.tool()` 注册同名函数**会触发 `Tool already exists` 警告且行为异常。

## 📚 Skill 层

`skill/duduexcel/` 提供 Agent 方法论（非 MCP 工具）：

```
SKILL.md                    # 工具路由表 + 5 条铁律 + 公式约束
references/style.md         # 配色语义、数字格式码、结构规范
references/formulas.md      # 函数白名单、_xlfn. 前缀、7 类错误速查
references/charts.md        # 图表做法与保真度
```

> 合规说明：Anthropic 官方 xlsx skill 为 Proprietary（禁止衍生作品）。
> 本 Skill 仅借鉴其**工程思想**（验证优先、截断诚实性、外链熔断），
> 文字与代码均为独立撰写。

## 🗺️ 路线图

| 阶段 | 状态 |
|---|---|
| M1 骨架 + 安全读写 | ✅ |
| M2 服务端分析 + token 自报 | ✅ |
| M3 公式重算 + 外链熔断 | ✅ |
| M4 中文样式 + 图表 | ✅ |
| M5 Skill 层 | ✅ |
| 透视表 / 条件格式 / 多表关联 | 待做 |

## License

MIT
