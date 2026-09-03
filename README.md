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

## ✨ 能力（13 个工具）

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

## ⚠️ 已知限制（诚实清单）

- **公式重算需要 LibreOffice**：未安装时 `recalculate` 明确降级并给出安装指引（不静默假装成功）。
  此时写入的公式无缓存值，`read_range` 读回 `None` —— 这是预期行为，非数据丢失。
- **不支持透视表与条件格式**（图表已支持）
- **`.xlsm` 宏**：读取保留 VBA，写入不保证
- 单次写入上限 **10 万单元格**
- 小表上 `_meta.tokens_saved` 节省不明显属正常（省 token 的收益随表增大而放大）

## 🧪 测试

```bash
python tests/test_smoke.py   # M1 读写与分页（17 项）
python tests/test_m2.py      # M2 服务端分析（25 项）
python tests/test_m34.py     # M3 重算降级 + M4 中文样式与图表（19 项）
```

全部 **61 项通过**，另有 14 项 MCP stdio 协议端到端验证。

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
