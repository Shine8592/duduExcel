# 📊 duduExcel

[![CI](https://github.com/Shine8592/duduExcel/actions/workflows/ci.yml/badge.svg)](https://github.com/Shine8592/duduExcel/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/duduexcel.svg)](https://pypi.org/project/duduexcel/)
[![Python](https://img.shields.io/pypi/pyversions/duduexcel.svg)](https://pypi.org/project/duduexcel/)
[![Downloads](https://img.shields.io/pypi/dm/duduexcel.svg)](https://pypi.org/project/duduexcel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-purple.svg)](https://modelcontextprotocol.io)

面向 AI Agent 的 Excel MCP 服务 —— **上下文高效 + 安全读写 + 中文场景 + 公式验证**。

**[English](README_EN.md) | 中文**

> 其他 Excel MCP 会把"已取消"的行和"生效中"的行读成一模一样的字符串，
> duduExcel 保留删除线 —— 因为**作者留下的格式，本身就是语义**。

```bash
pip install "duduexcel[analysis]"    # 或：uvx duduexcel
```

```
Agent ──MCP(stdio)──► duduExcel ──► openpyxl/pandas ──► .xlsx
```

---

## 为什么又一个 Excel MCP？

调研 GitHub 上 10+ 个同类项目后发现，它们分裂成两个对立阵营，没人同时做到"能写"和"不炸上下文"：

| 阵营 | 代表 | 能写 | 上下文友好 | 问题 |
|---|---|---|---|---|
| 读写全能型 | haris-musa、knorq | ✅ | ❌ | 整表读进上下文就爆了 |
| 原子分析型 | jwadow、jdatamunch | ❌ 只读 | ✅ | 改不了文件 |
| 方法论型 | Anthropic 官方 xlsx skill | ✅（靠写代码） | ✅ | 靠 Agent 自己写 Python，不稳定 |

**duduExcel：三者合一** —— 服务端原子分析（省 token）+ 完整读写 + 内置 Skill。

## ✨ 能力（19 个工具）

| 层 | 工具 | 说明 |
|---|---|---|
| **探查** | `workbook_info` | 表清单、行列数、合并单元格数、隐藏行列数、内嵌图片数、文件大小 |
| | `sheet_profile` | 列画像：类型/空值率/唯一数/Top值/统计（一次替代十几次调用） |
| **分析** | `filter_count` | 条件计数（14 种运算符），只回传数字 |
| | `aggregate` | sum/mean/count/… 支持分组与过滤 |
| | `top_n` | 排行榜，只回传前 N 行 |
| **读写** | `read_range` | 分页读取（默认 limit=200）；含格式语义与隐藏处理 |
| | `write_cells` | **批量**写入，一次调用完成；`=` 开头即公式 |
| **验证** ⭐ | `recalculate` | 公式重算 + 外链熔断（差异化，竞品多不支持） |
| | `scan_formula_errors` | 扫描 7 类公式错误 |
| **中文** ⭐ | `apply_chinese_style` | 微软雅黑表头、中文列宽自适应、冻结首行、细边框 |
| | `set_number_format` | `¥#,##0` / `0.0%` / `0.0x` 等内置格式 |
| **图表** ⭐ | `add_chart` | bar/line/pie/scatter |
| **高级** | `create_pivot` | 静态透视汇总表（诚实标注不可交互） |
| | `add_conditional_format` | 数据条/色阶/阈值高亮/区间/重复值 |
| | `list_conditional_formats` | 读取已有条件格式（写入+读取闭环） |
| | `compare_sheets` | 两表按关键列比对，只回差异摘要 |
| | `join_sheets` | 两表关联（left/right/inner/outer），只回前 N 行 |
| | `list_images` | 列出内嵌图片（零依赖扫描 `xl/media/`） |
| **安全** | `revert_last_write` | 回滚最近一次写入 |

⭐ = 差异化能力

## 🔑 汲取的设计（附来源）

| 设计点 | 来源 | 落地 |
|---|---|---|
| 一次调用替代 N 次试探 | jwadow `get_data_profile` | `sheet_profile` |
| 服务端原子操作（results, not rows） | jwadow | `filter_count`/`aggregate`/`top_n` |
| `_meta.tokens_saved` 自报节省 | jdatamunch | 每个分析工具的 `_meta` |
| 结果附 Excel 公式（可复现） | jwadow | `filter_count`/`aggregate`/`top_n` |
| TSV 输出（便于粘回 Excel） | jwadow | `aggregate`/`top_n` |
| 批量接口，禁止循环调用 | knorq | `write_cells` |
| 路径白名单 + 拒目录穿越 | haris-musa | `DUDU_EXCEL_ROOT` |
| 外链熔断、诚实截断、交付前验证 | 官方 `recalc.py` | `recalculate` |
| 格式语义、隐藏处理、内嵌图片、原子保存 | excel-vision-mcp | `read_range` 标记 / `list_images` / 原子保存 |

## 🚀 安装与挂载

```bash
pip install "duduexcel[analysis]"     # analysis 启用 pandas 分析工具
uvx duduexcel                          # 免安装运行
```

opencode（`~/.config/opencode/opencode.jsonc`）：

```jsonc
"duduexcel": {
  "type": "local",
  "command": ["python", "-u", "-m", "duduexcel"],
  "enabled": true
}
```

Claude Desktop / Cursor / Cline：

```json
{ "mcpServers": { "duduexcel": { "command": "uvx", "args": ["duduexcel"] } } }
```

**改完配置需重启客户端生效。**

## 🔄 第二轮调研补齐的盲区

调研 `VOYAGER-Inc/excel-vision-mcp` 后发现 openpyxl 生态的普遍盲区（duduExcel 原本也有）：
**内嵌图片、格式语义、隐藏行列全部丢失**。

```
修复前  [["旧版导出","已取消"], ["司机点名","待审阅"]]
修复后  A3: 旧版导出 [S] | B3: 已取消 [S]        ← [S]=删除线（已取消）
        A4: 司机点名 | B4: 待审阅 [HL:yellow]    ← [HL:]=黄底（待审阅）
```

已补齐：格式语义标记、隐藏内容智能处理、内嵌图片清单、原子保存、多目录沙箱。

## 🔒 安全

- **多目录沙箱**：`DUDU_EXCEL_ROOT` 支持多个目录（Windows `;`、其他 `:`），拒绝绝对路径与 `..` 穿越
- **原子保存**：写临时文件成功后再替换，失败的写入永不损坏原文件
- **备份与回滚**：写前 `.bak`，异常自动还原，`revert_last_write` 可撤销
- **外链熔断**：拒绝不可逆的重算，除非显式 `force=true`
- **本地优先**：stdio 传输，文件不出本机

## ⚠️ 已知限制（诚实清单）

- **不支持可交互透视表**：`create_pivot` 生成静态汇总表（数值已验证，如华东 365 / 华南 210），
  但 openpyxl 无法创建真 PivotTable（实测 `ws._pivots` 为空），故不可点击交互。
- ~~条件格式只能写入不能读取~~ —— **已更正**：实测**可以完整读回**
  （作用区域/类型/运算符/阈值/填充色/优先级，4 条规则全部读回）。
  此前我照抄竞品 knorq 的 Known Limitations 却未亲自验证，这是错误的。
- **`.xlsm` 宏**：读取保留 VBA，写入不保证
- 单次写入上限 **10 万单元格**
- **重算已在本机打通验证**（LibreOffice 26.8）：实测 **4 秒**完成，
  `=SUM(销售!C2:C7)` 正确算出 **575** 并落盘。未装时明确降级，绝不静默假装成功。
  - ⚠️ **中文路径坑（已修复）**：LibreOffice 在中文路径下原地覆盖会失败
    （`SfxBaseModel::impl_store failed: 0x4c0c`）。改为输出到纯 ASCII 临时目录再移回，
    因此 `E:\工作类\研发\` 这类路径也能正常重算。

## 🧪 测试

```bash
python tests/test_smoke.py          # M1 读写与分页
python tests/test_m2.py             # M2 服务端分析
python tests/test_m34.py            # M3 重算降级 + M4 中文样式与图表
python tests/test_m6.py             # M6 透视表/条件格式/多表关联
python tests/test_edge_cases.py     # 边界回归（防 BUG 复发）
python tests/test_m7.py             # M7 格式语义/隐藏/图片/原子保存
python tests/test_m8_recalc.py      # M8 真实重算（无 LibreOffice 自动跳过）
python tests/test_m9_conditional.py # M9 条件格式读取
```

9 个测试文件**全部通过**，另有 **20 项 MCP 端到端验证**（覆盖全部 19 个工具）。

## 🧠 踩过的坑

1. **openpyxl `read_only=True` 不加载行列维度** —— 隐藏检测会静默失效，改为解析 sheet XML
2. **mcp 2.x 丢弃 `str \| None`（PEP 604）标注的参数** —— 传了却不生效，改用 `Optional[str]`
3. **图表 `Reference(range_string="B2:B5")` 要求 `表名!A1:B2`** —— 改用 `range_boundaries()`
4. **重复 `@mcp.tool()` 注册同名函数** → `Tool already exists` 且行为异常
5. **内部函数误加 `@mcp.tool()`**（`_as_bool` 曾成伪工具，污染列表）
6. **构造结果后漏 `rules.append(item)`** —— 遍历正常却无结果，属静默 BUG

## 📚 Skill 层

```
skill/duduexcel/
├── SKILL.md                  # 工具路由 + 5 条铁律 + 公式约束
└── references/
    ├── style.md              # 财务配色语义、数字格式码
    ├── formulas.md           # 函数白名单、_xlfn. 前缀、7 类错误
    └── charts.md             # 图表做法与保真度
```

> 合规：Anthropic 官方 xlsx skill 为 **Proprietary**（禁止衍生作品）。
> 本 Skill 仅借鉴其**工程思想**，文字与代码均独立撰写。

## 🗺️ 路线图

- [x] M1 核心 + 安全读写
- [x] M2 服务端分析 + token 自报
- [x] M3 公式重算 + 外链熔断
- [x] M4 中文样式 + 图表
- [x] M5 Skill 层
- [x] M6 透视表 / 条件格式 / 多表关联
- [x] M7 格式语义 / 隐藏 / 图片 / 原子保存
- [x] M8 真实重算闭环（含中文路径修复）
- [x] M9 条件格式读取
- [ ] 可交互透视表（受 openpyxl 限制）

## License

MIT © Shine8592
