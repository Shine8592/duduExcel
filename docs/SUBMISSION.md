# 📢 Submission Copy for MCP Directories

Ready-to-paste descriptions for submitting duduExcel to MCP directories.
Targets: **glama.ai**, **mcp.so**, **Smithery**, **PulseMCP**, **awesome-mcp-servers** (PR),
**mcpservers.org**, **LobeHub MCP Market**.

---

## Short description (≤120 chars)

> Excel MCP server for AI agents — context-efficient, safe writes, format semantics, formula recalc, Chinese-ready.

*(119 characters)*

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

18 TOOLS
  Inspect   : workbook_info, sheet_profile
  Analyze   : filter_count, aggregate (incl. group_by), top_n
  Read/Write: read_range, write_cells, revert_last_write
  Verify    : recalculate, scan_formula_errors
  Chinese   : apply_chinese_style, set_number_format
  Charts    : add_chart (bar/line/pie/scatter)
  Advanced  : create_pivot, add_conditional_format, compare_sheets, join_sheets, list_images

INSTALL
  pip install "duduexcel[analysis]"
  # or: uvx duduexcel

  { "mcpServers": { "duduexcel": { "command": "uvx", "args": ["duduexcel"] } } }

REQUIREMENTS
  Python 3.10+. Formula recalculation additionally needs LibreOffice; without it the tool degrades
  explicitly with install instructions — it never silently pretends to succeed.

HONEST LIMITATIONS
  • Interactive PivotTables are not supported (create_pivot yields a static summary table;
    openpyxl cannot create real PivotTable objects).
  • Conditional formatting can be written but not read.
  • .xlsm macros preserved on read, not guaranteed on write.
  • Max 100,000 cells per write.

LINKS
  GitHub: https://github.com/Shine8592/duduExcel
  License: MIT
```

---

## awesome-mcp-servers PR entry (markdown)

```markdown
- [duduExcel](https://github.com/Shine8592/duduExcel) - 📊 Excel MCP server: context-efficient
  (server-side aggregation + pagination), format semantics (strikethrough/highlight convey intent,
  not just values), hidden-content handling, embedded-image listing, formula recalculation with an
  external-link circuit breaker, atomic/backup/rollback-safe writes, and first-class Chinese support.
  [`excel`][`xlsx`][`spreadsheet`] `python` `mit`
```

---

## Category / tag suggestions

- **Category**: Data & Analytics · Productivity · File Systems
- **Tags**: `excel` `xlsx` `spreadsheet` `mcp` `model-context-protocol` `ai-agent` `llm`
  `openpyxl` `pandas` `automation` `chinese` `local-first`

---

## One-line pitch (for social posts)

> Other Excel MCP servers flatten a cancelled row and an active row into identical strings.
> duduExcel keeps the strikethrough — because the author's formatting *is* the meaning.
