# 📊 duduExcel

**Excel MCP server built for AI agents** — context-efficient, safe to write, and it understands what the author *meant*, not just cell values.

```
Agent ──MCP (stdio)──► duduExcel ──► openpyxl / pandas ──► .xlsx
```

*Read this in [中文](README.md)*

---

## Why another Excel MCP?

Most Excel MCP servers make you choose between two things:

| Approach | Examples | Can write? | Context-friendly? |
|---|---|---|---|
| Full read/write | haris-musa, knorq | ✅ | ❌ dumps whole sheets |
| Atomic analysis | jwadow, jdatamunch | ❌ read-only | ✅ |
| Methodology only | Anthropic's xlsx skill | ✅ (via generated code) | ✅ but relies on the model writing Python |

**duduExcel does all three**: server-side aggregation (cheap), full read/write (capable), and a bundled Skill (reliable).

### And it sees what others silently drop

> A strikethrough row means *cancelled*. A yellow-highlighted cell means *needs review*. A pasted flowchart in `B12` holds half the spec.
> **Every other Excel MCP flattens all of that into plain strings** — so a cancelled row looks identical to an active one.

```
# Other MCP servers
[["Legacy export", "Cancelled"], ["Driver roll call", "Needs review"]]

# duduExcel
A3: Legacy export [S] | B3: Cancelled [S]        ← [S] = strikethrough
A4: Driver roll call  | B4: Needs review [HL:yellow]  ← [HL:] = highlight
```

---

## ✨ 18 Tools

| Layer | Tools |
|---|---|
| **Inspect** | `workbook_info` · `sheet_profile` |
| **Analyze** (server-side) | `filter_count` · `aggregate` · `top_n` |
| **Read / Write** | `read_range` · `write_cells` · `revert_last_write` |
| **Verify** ⭐ | `recalculate` · `scan_formula_errors` |
| **Chinese** ⭐ | `apply_chinese_style` · `set_number_format` |
| **Charts** ⭐ | `add_chart` |
| **Advanced** | `create_pivot` · `add_conditional_format` · `compare_sheets` · `join_sheets` · `list_images` |

⭐ = capabilities competitors commonly lack

### Highlights

- **Context-efficient by design** — `read_range` paginates (default `limit=200`) and reports `truncated` honestly. Analysis runs server-side: you get `{"sum": 575}`, not 5,000 rows. Every response carries `_meta.tokens_saved`.
- **Format semantics** — `[B]` bold · `[I]` italic · `[S]` strikethrough (cancelled) · `[HL:color]` highlight (needs review) · `[C:color]` font color · `[M]` merged · `[HIDDEN-REF]` hidden-but-referenced. **Markers are only attached when formatting is actually used**, so plain sheets cost zero extra tokens.
- **Hidden content, handled** — hidden rows/cols are skipped by default (an author hiding them signals they're not for review), **except** cells referenced by visible formulas, which are kept and tagged `[HIDDEN-REF]`. Skipped counts are always reported — nothing disappears silently.
- **Embedded images** — `list_images` scans `xl/media/` with zero extra dependencies, so flowcharts pasted into cells no longer vanish.
- **Formula recalculation** — via LibreOffice, with **external-link circuit breaker**: if a workbook has external links whose cached values were stripped by openpyxl, recalculation would turn them into `#NAME?` **and delete the links permanently**. duduExcel refuses by default and tells you why; pass `force=true` only if you accept the loss.
- **Safe writes** — atomic save (temp file → swap), `.bak` backup before every write, auto-rollback on failure, `revert_last_write` to undo.
- **Chinese-first** — Microsoft YaHei headings, CJK-aware column autofit (2 units per CJK char), `¥#,##0` currency, `0.0%`, `0.0x` multiples.

---

## 🚀 Install

```bash
pip install "duduexcel[analysis]"     # analysis extra enables pandas-backed tools
# or run without installing:
uvx duduexcel
```

**Requirements:** Python 3.10+. Formula recalculation (`recalculate`) additionally needs [LibreOffice](https://www.libreoffice.org/) installed; without it the tool degrades explicitly with install instructions — it never silently pretends to succeed.

---

## 🔌 MCP Configuration

### opencode (`~/.config/opencode/opencode.jsonc`)

```jsonc
"duduexcel": {
  "type": "local",
  "command": ["python", "-u", "-m", "duduexcel"],
  "enabled": true
}
```

### Claude Desktop / Cursor / Cline

```json
{
  "mcpServers": {
    "duduexcel": {
      "command": "uvx",
      "args": ["duduexcel"]
    }
  }
}
```

Restart your client after editing the config.

---

## 🔒 Security

- **Path sandbox** — set `DUDU_EXCEL_ROOT` to one or more directories (`;` on Windows, `:` elsewhere); absolute paths and `..` traversal outside them are rejected.
- **Atomic save** — a failed write can never corrupt your original file.
- **Auto backup / rollback** — `.bak` before every write, auto-restore on exception, `revert_last_write` to undo.
- **External-link circuit breaker** — refuses irreversible recalculation unless you force it.
- **Local-first** — stdio transport; your files never leave the machine.

---

## ⚠️ Known Limitations (honest list)

- **Interactive PivotTables are not supported.** `create_pivot` produces a **static summary table** (group-and-aggregate written back) — numerically equivalent and verified correct (e.g. East 365 / South 210), but not clickable. openpyxl cannot create real PivotTable objects (`ws._pivots` is empty).
- ~~Conditional formatting can be written, not read~~ — **corrected**: it *can* be read. `list_conditional_formats` returns range, type, operator, thresholds, fill color, and priority (verified: 4/4 rules read back). This limitation was mistakenly copied from a competitor's README without verification.
- **`.xlsm` macros** are preserved on read; not guaranteed on write.
- Max **100,000 cells** per write.
- `recalculate` needs LibreOffice (verified working on 26.8, ~4s for a typical sheet).

---

## 🧪 Tests

```bash
python tests/test_smoke.py       # read/write, pagination, safety
python tests/test_m2.py          # server-side analysis
python tests/test_m34.py         # recalc degradation + Chinese styling + charts
python tests/test_m6.py          # pivot / conditional format / multi-sheet
python tests/test_edge_cases.py  # regression guards for past bugs
python tests/test_m7.py          # format semantics / hidden / images / atomic save
python tests/test_m8_recalc.py   # real recalculation (skips if no LibreOffice)
```

---

## 📚 Bundled Agent Skill

`skill/duduexcel/` ships methodology the agent can auto-load:

```
SKILL.md                  # tool routing + 5 hard rules + formula constraints
references/style.md       # financial-model color semantics, number formats
references/formulas.md    # function allowlist, _xlfn. prefixes, 7 error types
references/charts.md      # chart recipes and fidelity caveats
```

> Anthropic's official xlsx skill is **Proprietary** (derivative works prohibited). duduExcel borrows only its *engineering ideas* (validate-before-deliver, honest truncation, external-link breaker); all text and code are written independently.

---

## 🗺️ Roadmap

- [x] M1 core + safe read/write
- [x] M2 server-side analysis + token self-reporting
- [x] M3 formula recalculation + external-link breaker
- [x] M4 Chinese styling + charts
- [x] M5 Skill layer
- [x] M6 pivot / conditional format / multi-sheet joins
- [x] M7 format semantics / hidden content / embedded images / atomic save
- [x] Conditional-formatting reading (`list_conditional_formats`)
- [ ] Interactive PivotTables (blocked by openpyxl)

---

## License

MIT © Shine8592
