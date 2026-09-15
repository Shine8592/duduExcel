# 🚀 MCP 目录提交执行清单

> 目标：让 duduExcel 从"0 曝光"变成"可被陌生人发现"。
> 状态会随进度更新。

---

## ✅ 已完成

| # | 渠道 | 状态 | 说明 |
|---|---|---|---|
| 1 | **glama.ai** | ✅ **已收录** | https://glama.ai/mcp/servers/Shine8592/duduExcel |
| 2 | **mcpservers.org** | ⏳ **审核中（约 12h）** | 已提交，站点提示 12 小时审核 |
| 3 | PyPI | ✅ 已发布 v0.3.0 | https://pypi.org/project/duduexcel/ |
| 4 | GitHub Release | ✅ v0.2.0 → v0.3.0 | 4 个 release |
| 5 | 仓库 topics | ✅ 12 个 | 让 GitHub 搜索可见 |

> ⚠️ **mcp.so 不做**：该站提交需付费。已找到免费替代（见下方 P1）。

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
| CI | ✅ 6 matrix jobs green (Ubuntu + Windows × 3.10/3.11/3.12) |
| Tests | 9 files, 110+ assertions |

**Short description（108 字符）**：
```
Excel MCP: context-efficient, safe writes, format semantics, interactive PivotTables, formula recalc.
```

**One-line pitch**：
```
Other Excel MCP servers flatten a cancelled row and an active row into identical strings.
duduExcel keeps the strikethrough — because the author's formatting IS the meaning.
```

---

## 🟡 P1：免费目录（全部免费，建议逐个提交）

> mcpservers.org 和 PulseMCP 都会**自动从 GitHub 抓取**，填仓库地址即可。

| 优先级 | 站点 | 提交入口 | 状态 | 备注 |
|---|---|---|---|---|
| ~~1~~ | ~~mcpservers.org~~ | — | ✅ 已提交 | 等 12h 审核 |
| **2** | **LobeHub MCP** | https://lobehub.com/mcp → Publish | ⬜ 待提交 | 页面有 Publish 入口，中文友好 |
| **3** | **Smithery** | https://smithery.ai → Publish | ⬜ 待提交 | 页面有 Publish 入口（已被 Arcade.dev 收购）|
| **4** | **ModelScope MCP 广场** | https://modelscope.cn/mcp | ⬜ 待提交 | 国内流量，需登录；中文 README 占优 |
| ⏸️ | **PulseMCP** | https://www.pulsemcp.com/submit | ⏸️ **站点暂停** | 官方公告：「submissions still paused while we rework how we ingest listings」——**等它重开再提交** |

---

## 🔴 P0：awesome-mcp-servers（95000★，审核慢但长期有效）

⚠️ 该仓库有 **2449 个开放 PR**，审核慢；但收录后是长期被动流量。

### 步骤

1. **Fork** `punkpeye/awesome-mcp-servers`
2. 在 README.md 找到 **📊 Data & Analytics** 小节
3. 在列表**末尾**插入：

```markdown
- [duduExcel](https://github.com/Shine8592/duduExcel) - 📊 Excel MCP server: context-efficient (server-side aggregation + pagination), format semantics (strikethrough/highlight convey author intent, not just values), hidden-content handling, embedded-image listing, **interactive PivotTables** via OOXML injection, formula recalculation with an external-link circuit breaker, atomic/backup/rollback-safe writes, and first-class Chinese support. 20 tools. `python` `mit`
```

4. **分支名**：`add-duduexcel`
5. **PR 标题**：
```
Add duduExcel - Excel MCP server with format semantics, interactive PivotTables and formula recalc
```
6. **PR 描述**：
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

---

## 🟢 P2：让 star 涨得更快（软性动作）

1. **README 放真实效果图/GIF** — 目前全文字，用户扫 3 秒决定成败
   - 建议录：写公式 → recalculate → 显示 575 → 生成透视表
   - 工具：ScreenToGif（Windows 免费）

2. **自己先用它干几件真活** — 用 duduExcel 处理真实表格并分享过程
   - "0 issue" 说明还没有真实用户，先用起来才会有反馈

3. **demo 已就绪** — `examples/demo.xlsx` + `examples/README.md`，访问者 30 秒可跑通 ✅

---

## 🔒 安全待办（重要）

对话中泄露过两个 token，**都仍然有效**（已验证）：

| Token | 状态 | 撤销方式 |
|---|---|---|
| GitHub PAT `ghp_zBiY...` | ⚠️ 仍有效 | https://github.com/settings/tokens → 删除该 token |
| PyPI token `pypi-AgEI...` | ⚠️ 仍有效 | https://pypi.org/manage/account/token/ → revoke |

> 这两处都**只能由你在网页操作**（PyPI 无 token 管理 API；GitHub PAT 撤销也需网页）。
> 我在对话里能调用的 API 权限不包括管理 token 本身。

---

## 📊 预期效果

| 动作 | 预期 |
|---|---|
| glama.ai 收录 | ✅ 已完成，几天内开始有陌生访问 |
| mcpservers.org / PulseMCP 收录 | 附带流量入口 |
| awesome-mcp-servers 合并 | 长期被动流量（95000★ 仓库引用）|
| 一条社交媒体帖 | 若内容好，可能带来前 10-50 star |

**第一个真实 issue 出现时，才算真正"活"了** —— 那是比 star 更硬的指标。
