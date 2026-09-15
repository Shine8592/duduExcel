# 🚀 MCP 目录提交执行清单

> 目标：让 duduExcel 从"0 曝光"变成"可被陌生人发现"。
> 按优先级从上到下执行，**P0 三件做完就有真实流量**。

---

## 📌 项目现成信息（填表时直接复制）

| 字段 | 值 |
|---|---|
| Name | `duduExcel` |
| Repository | `https://github.com/Shine8592/duduExcel` |
| PyPI | `https://pypi.org/project/duduexcel/` |
| Install | `uvx duduexcel` 或 `pip install "duduexcel[analysis]"` |
| Language | Python |
| License | MIT |
| Version | 0.3.0 |
| Tools | 20 |
| CI | ✅ 6 matrix jobs green (Ubuntu+Windows × 3.10/3.11/3.12) |
| Tests | 9 files, 110+ assertions |

**Short description（108 字符，通用）**：
```
Excel MCP: context-efficient, safe writes, format semantics, interactive PivotTables, formula recalc.
```

**One-line pitch**：
```
Other Excel MCP servers flatten a cancelled row and an active row into identical strings.
duduExcel keeps the strikethrough — because the author's formatting IS the meaning.
```

---

## 🔴 P0-1：glama.ai（流量最大，首选）

1. 打开 https://glama.ai/mcp/servers
2. 右上角 **"Submit a Server"** / **"Add Server"**
3. 选择 **"GitHub Repository"** 方式（不是手动填）
4. 填入：`https://github.com/Shine8592/duduExcel`
5. 提交后它会自动从仓库 README 抓取描述（我们已备好英文 README，含徽章）
6. 若要求补充分类：选 **Data & Analytics**；Tags 填 `excel, xlsx, spreadsheet, mcp, ai-agent`

> 💡 glama 会自动跑一遍 MCP server，验证 `tools/list` 能否工作。我们的 server 已通过 CI 验证，应无问题。

---

## 🔴 P0-2：mcp.so

1. 打开 https://mcp.so/submit
2. 表单：
   - **Name**: `duduExcel`
   - **URL/GitHub**: `https://github.com/Shine8592/duduExcel`
   - **Description**: 用上面的 short description
   - **Category**: Data & Analytics
   - **Install**: `uvx duduexcel`
3. 提交（可能需登录 GitHub 账号）

---

## 🔴 P0-3：awesome-mcp-servers（95000★，PR 多，但收录后长期有效）

⚠️ 该仓库有 **2449 个开放 PR**，审核慢；但收录后是长期被动流量。

### 步骤

1. **Fork** `punkpeye/awesome-mcp-servers`
2. 在 README.md 中找到 **📊 Data & Analytics** 分类小节
3. 在该小节列表**末尾**插入这一行：

```markdown
- [duduExcel](https://github.com/Shine8592/duduExcel) - 📊 Excel MCP server: context-efficient (server-side aggregation + pagination), format semantics (strikethrough/highlight convey author intent, not just values), hidden-content handling, embedded-image listing, **interactive PivotTables** via OOXML injection, formula recalculation with an external-link circuit breaker, atomic/backup/rollback-safe writes, and first-class Chinese support. 20 tools. `python` `mit`
```

4. **分支名**：`add-duduexcel`
5. **PR 标题**：
```
Add duduExcel - Excel MCP server with format semantics, interactive PivotTables and formula recalc
```
6. **PR 描述**（直接复制）：
```markdown
## What is duduExcel?

An Excel MCP server for AI agents that is context-efficient, safe to write, and reads the
author's *intent* — not just cell values.

### Why it's different

Most Excel MCP servers fall into two camps: capable but context-heavy, or context-efficient
but read-only. duduExcel does both, and adds things others silently drop:

- **Format semantics** — a strikethrough row (cancelled) is not the same as an active one.
  Returns `[B] [I] [S] [HL:color] [C:color] [M]` markers, attached only when formatting is
  actually used (plain sheets cost zero extra tokens).
- **Hidden content** — hidden rows/cols skipped by default, except cells referenced by visible
  formulas (kept and tagged `[HIDDEN-REF]`); skip counts always reported.
- **Interactive PivotTables** — openpyxl cannot create PivotTables, so duduExcel injects the
  OOXML pivot cache + layout parts itself. Validated by LibreOffice open/re-save (all 5 parts
  survive). A static summary flavour is also available.
- **Formula recalculation** — with an external-link circuit breaker: refuses irreversible
  recalculation unless explicitly forced.
- **Safe writes** — atomic save, `.bak` backup, auto-rollback, `revert_last_write`,
  multi-directory path sandbox.
- **Chinese-first** — YaHei headings, CJK-aware autofit, `¥#,##0`.

### Quality signals

- ✅ CI green: Ubuntu + Windows × Python 3.10/3.11/3.12 (6 jobs)
- ✅ 9 test files, 110+ assertions (incl. a real LibreOffice validation)
- ✅ Published on PyPI: `pip install "duduexcel[analysis]"`
- ✅ 20 tools, MIT licensed, bilingual README

### Limitations (stated honestly)

- An interactive PivotTable is dropped if the file is later saved with openpyxl — generate last.
- No pivot field grouping / calculated fields / slicers / data model.
- `recalculate` requires LibreOffice; without it the tool degrades explicitly, never silently.
```

7. 提交 PR

---

## 🟡 P1：其他目录（次优先，批量提交）

| 站点 | 提交入口 | 备注 |
|---|---|---|
| **PulseMCP** | https://www.pulsemcp.com/submit | 自动抓 GitHub |
| **mcpservers.org** | https://mcpservers.org/submit | 表单 |
| **LobeHub MCP** | https://lobehub.com/mcp/submit | 表单，中文友好 |
| **Smithery** | https://smithery.ai/new | 需 Smithery CLI，略麻烦 |
| **ModelScope MCP 广场** | https://modelscope.cn/mcp | 国内流量，中文 README 有优势 |

---

## 🟢 P2：让 star 涨得更快（软性动作）

1. **在 README 顶部放一张真实效果图/GIF**
   - 目前 README 全是文字。用户扫 3 秒决定是否 star。
   - 建议录一个：写公式 → recalculate → 显示 575 → 生成透视表
   - 工具：ScreenToGif（Windows 免费）

2. **自己先用它干几件真活**
   - 用 duduExcel 处理你自己的表格，把过程发一条推/掘金/少数派
   - "0 issue" 说明还没真实用户，先用起来才会有反馈

3. **给仓库加一个 demo 文件**
   - `examples/demo.xlsx` + `examples/README.md`，让访问者 30 秒能跑通

---

## ⚠️ 提交前最后检查（已完成 ✅）

- [x] README 顶部有 CI / PyPI 徽章
- [x] 英文 README（国际目录必需）
- [x] LICENSE 存在（MIT）
- [x] CHANGELOG 完整
- [x] PyPI 可安装
- [x] CI 全绿
- [x] 仓库 topics 已设（12 个）
- [x] GitHub Release 已建（v0.2.0 / v0.2.1 / v0.2.2 / v0.3.0）

---

## 📊 预期效果

| 动作 | 预期 |
|---|---|
| glama.ai 收录 | 几天内开始有陌生访问 |
| mcp.so 收录 | 同上 |
| awesome-mcp-servers 合并 | 长期被动流量（95000★ 仓库的引用） |
| 一条社交媒体帖 | 若内容好，可能带来前 10-50 star |

**第一个真实 issue 出现时，才算真正"活"了** —— 那是比 star 更硬的指标。
