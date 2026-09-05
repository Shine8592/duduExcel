# Changelog

All notable changes to duduExcel. This project uses [Semantic Versioning](https://semver.org/).

## [0.2.0] — 2026-09-05

### Added

#### M7 — Format semantics, hidden content, embedded images (inspired by VOYAGER-Inc/excel-vision-mcp)
- **Format semantics markers** in `read_range`: `[B]` bold, `[I]` italic, `[S]` strikethrough (cancelled),
  `[HL:color]` highlight (needs review), `[C:color]` font color, `[M]` merged cell
  - Markers are attached **only when formatting is actually used** — plain sheets cost zero extra tokens
  - Includes `marker_legend` and a `formatted_view` for quick reading
- **Hidden content handling**: hidden rows/columns skipped by default; cells referenced by visible
  formulas are kept and tagged `[HIDDEN-REF]`; skip counts always reported (never silent)
- **`list_images`**: lists embedded images via zero-dependency `xl/media/` scan (count, filename, size, dimensions)
- `workbook_info` now reports merged-cell count, hidden row/col counts, and embedded image count
- **Atomic save**: write to temp → swap, so a failed write can never corrupt the original
- **Multi-directory sandbox**: `DUDU_EXCEL_ROOT` accepts several paths (`;` on Windows, `:` elsewhere)

#### M6 — Advanced operations
- `create_pivot` (static summary table; honestly labelled as non-interactive)
- `add_conditional_format` (data bar, color scale, greater/less/equal, between, duplicate)
- `compare_sheets` (keyed diff: only-in-left / only-in-right / value differences)
- `join_sheets` (left/right/inner/outer, returns first N rows only)

#### Packaging & docs
- `LICENSE` (MIT), `MANIFEST.in`, English `README_EN.md`
- GitHub Actions CI (Python 3.10–3.12 × Ubuntu/Windows) + build job
- `scripts/install_libreoffice.py`, `scripts/check_libreoffice.py`, `scripts/health_check.py`
- `docs/SUBMISSION.md` — ready-to-paste copy for MCP directories

### Fixed
- **Recalculation failed on Chinese paths** — LibreOffice cannot overwrite a file in place when the path
  contains non-ASCII characters (`SfxBaseModel::impl_store failed: 0x4c0c`).
  Now converts into an ASCII temp directory and moves the result back. Verified on `E:\工作类\研发\`.
- Non-zero soffice exit now surfaces stderr instead of only reporting "file not rewritten"
- **`_as_bool` internal helper was mistakenly decorated with `@mcp.tool()`**, registering it as a bogus
  19th tool and polluting `tools/list`. Removed the decorator (back to 18 tools) with a guard comment.
- `test_m34.py` adapts to environment: asserts real recalculation when LibreOffice exists,
  degradation path when it does not

## [0.1.0] — 2026-09-03

### Added

#### M1 — Core + safe read/write
- `workbook_info`, `read_range` (pagination + honest truncation), `write_cells` (bulk), `revert_last_write`
- Path sandbox (`DUDU_EXCEL_ROOT`), backup-before-write, auto rollback on failure

#### M2 — Server-side analysis
- `sheet_profile` (one call replaces a dozen probes), `filter_count`, `aggregate` (incl. `group_by`), `top_n`
- `_meta.tokens_saved` self-reporting, equivalent Excel formulas and TSV output on results

#### M3 — Formula verification
- `recalculate` with external-link circuit breaker and file-fingerprint silent-failure guard
- `scan_formula_errors` (7 Excel error types, honest truncation reporting)

#### M4 — Chinese scenarios & charts
- `apply_chinese_style` (YaHei headings, CJK-aware autofit, freeze panes), `set_number_format` (`¥#,##0`, `0.0%`, `0.0x`)
- `add_chart` (bar/line/pie/scatter)

#### M5 — Skill layer
- `SKILL.md` + `references/{style,formulas,charts}.md`

### Fixed
- Type-handling crashes found by boundary probing:
  1. `TypeError` when comparing text columns with `>`/`<` — now a readable error suggesting text operators
  2. `TypeError: bad operand type for unary -: 'str'` when aggregating text columns with `group_by` —
     replaced with a type-safe sort key

---

## Design lineage

Ideas borrowed (never code) from the ecosystem:

| Idea | Source |
|---|---|
| Server-side atomic operations ("results, not rows") | jwadow/mcp-excel |
| `_meta.tokens_saved`, destructive-action preflight | jgravelle/jdatamunch-mcp |
| Bulk tools, explicit "don't loop the single version" | knorq-ai/xlsx-mcp-server |
| Path sandbox (`EXCEL_FILES_PATH`) | haris-musa/excel-mcp-server |
| External-link breaker, honest truncation, validate-before-deliver | Anthropic official xlsx skill (Proprietary — engineering ideas only) |
| Format semantics, hidden-content handling, embedded images, atomic save | VOYAGER-Inc/excel-vision-mcp |
