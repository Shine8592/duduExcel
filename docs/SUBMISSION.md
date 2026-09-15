# 📢 Submission Copy for MCP Directories

Ready-to-paste descriptions for submitting **duduExcel** to MCP directories.

Targets: **glama.ai**, **mcp.so**, **PulseMCP**, **awesome-mcp-servers** (PR),
**mcpservers.org**, **LobeHub MCP Market**, **Smithery**.

> 更新于 v0.3.0：20 个工具、可交互透视表、条件格式读写、CI 全绿。

---

## Short description (≤120 chars)

> Excel MCP server for AI agents — context-efficient, safe writes, format semantics, interactive PivotTables, formula recalc, Chinese-ready.

*(134 chars — 若站点限制 120，用下面这版)*

> Excel MCP: context-efficient, safe writes, format semantics, interactive PivotTables, formula recalc.

*(108 characters)*

---

## Long description

```
duduExcel is an Excel MCP server built for AI agents. Most Excel MCP servers make you choose:
either they can write but dump entire sheets into your context, or they are context-efficient
but read-only. duduExcel does both — and sees what the others silently drop.

KEY CAPABILITIES

• Context-efficient by design
  read_range paginates (default limit=200) and reports truncation honestly. Analysis runs
  server-side: you get {"sum": 575}, not 5,000 rows. Every response carries _meta.tokens_saved.

• Format semantics — the author's intent, not just strings
  A strikethrough row means "cancelled"; a yellow cell means "needs review". Other MCP servers
  flatten both into plain text. duduExcel returns markers: [B] bold, [I] italic, [S] strikethrough,
  [HL:yellow] highlight, [C:red] font color, [M] merged. Markers appear only when formatting is
  actually used, so plain sheets cost zero extra tokens.

• Hidden content, handled intelligently
  Hidden rows/columns are skipped by default (an author hiding them signals they're not for review),
  EXCEPT cells referenced by visible formulas — those are kept and tagged [HIDDEN-REF], because
  their values drive results you can see. Skipped counts are always reported; nothing vanishes silently.

• Embedded images are not lost
  list_images scans xl/media/ with zero extra dependencies, so flowcharts pasted into cells no
  longer disappear. (Most Excel MCP servers silently drop every embedded image.)

• INTERACTIVE PivotTables — built by injecting OOXML parts
  openpyxl can only preserve PivotTables, never create them. duduExcel constructs the pivot cache
  and layout XML itself, so Excel opens a REAL PivotTable you can drag fields, expand/collapse
  and refresh. Validated by opening and re-saving with LibreOffice (all 5 pivot parts survive).
  A static summary flavour (create_pivot) is also available as a fast path.

• Formula recalculation with an external-link circuit breaker
  Via LibreOffice (~4s typical). If a workbook has external links whose cached values openpyxl
  stripped, recalculating turns them into #NAME? AND deletes the links permanently. duduExcel
  refuses by default and explains why; force=true only if you accept the loss.

• Safe writes
  Atomic save (temp → swap) so a failed write can never corrupt your file. .bak before every write,
  auto-rollback on failure, revert_last_write to undo. Path sandbox with multi-directory support.

• Chinese-first
  Microsoft YaHei headings, CJK-aware column autofit (2 units per CJK char), ¥#,##0 currency,
  0.0%, 0.0x multiples.

20 TOOLS
  Inspect   : workbook_info, sheet_profile
  Analyze   : filter_count, aggregate (incl. group_by), top_n
  Read/Write: read_range, write_cells, revert_last_write
  Verify    : recalculate, scan_formula_errors
  Chinese   : apply_chinese_style, set_number_format
  Charts    : add_chart (bar/line/pie/scatter)
  Advanced  : create_interactive_pivot, create_pivot, add_conditional_format,
              list_conditional_formats, compare_sheets, join_sheets, list_images

INSTALL
  pip install "duduexcel[analysis]"
  # or: uvx duduexcel

  { "mcpServers": { "duduexcel": { "command": "uvx", "args": ["duduexcel"] } } }

REQUIREMENTS
  Python 3.10+. Formula recalculation additionally needs LibreOffice; without it the tool degrades
  explicitly with install instructions — it never silently pretends to succeed.

HONEST LIMITATIONS
  • Saving again with openpyxl (incl. this server's own write tools) drops an interactive
    PivotTable — generate it LAST.
  • PivotTables: no field grouping, calculated fields/items, slicers, timelines, or data model.
  • .xlsm macros preserved on read, not guaranteed on write.
  • Max 100,000 cells per write.

CI: 6 matrix jobs green (Ubuntu + Windows x Python 3.10/3.11/3.12), plus a build job.
TESTS: 9 test files, 110+ assertions, including a real LibreOffice open/re-save check.

LINKS
  GitHub: https://github.com/Shine8592/duduExcel
  PyPI:   https://pypi.org/project/duduexcel/
  License: MIT
```

---

## awesome-mcp-servers PR entry (markdown)

```markdown
- [duduExcel](https://github.com/Shine8592/duduExcel) - 📊 Excel MCP server: context-efficient
  (server-side aggregation + pagination), format semantics (strikethrough/highlight convey intent,
  not just values), hidden-content handling, embedded-image listing, **interactive PivotTables**
  (built by injecting OOXML parts), formula recalculation with an external-link circuit breaker,
  atomic/backup/rollback-safe writes, and first-class Chinese support. 20 tools. `python` `mit`
```

---

## glama.ai / mcp.so / PulseMCP 表单填写建议

大多数目录要这些字段，直接复制：

| 字段 | 填写内容 |
|---|---|
| **Name** | `duduExcel` |
| **Description** | 用上面的 *Short description* |
| **Repository** | `https://github.com/Shine8592/duduExcel` |
| **Install command** | `uvx duduexcel` |
| **Language** | `Python` |
| **License** | `MIT` |
| **Category** | `Data & Analytics` / `Productivity` |
| **Tags** | `excel` `xlsx` `spreadsheet` `mcp` `ai-agent` `openpyxl` `pandas` `chinese` `pivot-table` |
| **Requires** | `Python 3.10+`; LibreOffice optional (only for `recalculate`) |
| **Env vars** | `DUDU_EXCEL_ROOT` (optional path sandbox), `UAM_EXTRA_SYS_PATH` (optional) |

---

## 一句话 pitch（社媒/README 用）

> Other Excel MCP servers flatten a cancelled row and an active row into identical strings.
> duduExcel keeps the strikethrough — because the author's formatting *is* the meaning.

---

## PR 提交步骤（awesome-mcp-servers）

1. Fork `punkpeye/awesome-mcp-servers`
2. 在 `README.md` 的 **📊 Data & Analytics**（或对应分类）小节末尾，插入上面的 markdown 条目
3. 分支名建议：`add-duduexcel`
4. PR 标题：`Add duduExcel - Excel MCP server with format semantics, interactive PivotTables and formula recalc`
5. PR 描述里贴上面的 *Long description* 首段，并注明 CI 绿 + PyPI 已发布

**注意**：提交前确认仓库 README 顶部已展示 CI / PyPI 徽章（已有），目录维护者会据此判断项目可信度。
