#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
核心 Excel 操作（纯函数，不依赖 MCP，便于单测）。

设计原则（汲取自调研）：
1. 上下文友好：任何"读"都必须能分页/截断，绝不把整表灌进 Agent 上下文
   （学 jwadow/mcp-excel 的 "results, not rows" 与 limit/offset）。
2. 批量优先：写入一次 load/save 完成多单元格（学 knorq 的 bulk 变体，
   它明确要求 "Use these instead of calling the single-target versions in a loop"）。
3. 诚实截断：截断时必须告知真实总数与被隐藏的行数
   （学 Anthropic 官方 recalc.py 的 "locations_truncated" 诚实性）。
4. 规避 openpyxl 已知坑（官方 xlsx skill 总结的六条）：
   - 只读场景用 read_only=True 提升大文件性能
   - 需要公式/值双读时分别以 data_only 开关加载
   - 写入用普通模式（read_only 工作簿不可写）
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 单次读取默认返回行数上限（防止 Agent 上下文被大表撑爆）
DEFAULT_READ_LIMIT = 200
# 单次写入的单元格数量上限（对齐 knorq 的"范围上限"思路，先设保守值）
MAX_WRITE_CELLS = 100_000


def _open(path: Path, data_only: bool = False, read_only: bool = False):
    """统一打开工作簿。xlsm 需保留 VBA，故按后缀决定 keep_vba。"""
    keep_vba = path.suffix.lower() == ".xlsm"
    return load_workbook(
        filename=str(path),
        data_only=data_only,
        read_only=read_only,
        keep_vba=keep_vba,
    )


def _sheet_names(path: Path) -> list[str]:
    wb = _open(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _pick_sheet(wb, sheet: str | None):
    """选择工作表；未指定用活动表；不存在时给出可读错误 + 可用列表。"""
    if sheet is None:
        return wb.active
    if sheet not in wb.sheetnames:
        raise KeyError(
            f"工作表 '{sheet}' 不存在。可用工作表：{', '.join(wb.sheetnames)}"
        )
    return wb[sheet]


# ---------------------------------------------------------------------------
# L1 探查：一次调用替代多次（学 jwadow 的 get_data_profile）
# ---------------------------------------------------------------------------
def inspect_workbook(file_path: Path) -> dict[str, Any]:
    """返回工作簿概览：工作表清单、每张表的行列数、文件大小。

    设计意图：让 Agent 用一次调用搞清楚文件结构，
    避免"先读全表再猜结构"这种把上下文吃光的做法。
    """
    size_bytes = file_path.stat().st_size
    wb = _open(file_path, read_only=True)
    try:
        sheets = []
        for ws in wb.worksheets:
            sheets.append(
                {
                    "name": ws.title,
                    "rows": ws.max_row or 0,
                    "columns": ws.max_column or 0,
                }
            )
        return {
            "file": str(file_path),
            "file_size_kb": round(size_bytes / 1024, 1),
            "active_sheet": wb.active.title if wb.active else None,
            "sheet_count": len(sheets),
            "sheets": sheets,
        }
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# L3 读取：分页 + 截断 + 诚实报告
# ---------------------------------------------------------------------------
def read_range(
    file_path: Path,
    sheet: str | None = None,
    cell_range: str | None = None,
    offset: int = 0,
    limit: int = DEFAULT_READ_LIMIT,
    include_formula: bool = True,
) -> dict[str, Any]:
    """读取单元格区域，支持分页与截断。

    - offset/limit：分页（limit 默认 200，防止上下文溢出）
    - 返回 total_rows / returned_rows / hidden_rows，截断时明确告知
    - include_formula=True 时同时给出公式与缓存值（学 knorq 的 read_cell 双字段）
    """
    if limit <= 0:
        raise ValueError("limit 必须大于 0")
    if offset < 0:
        raise ValueError("offset 不能为负数")

    wb_value = _open(file_path, data_only=True, read_only=True)
    try:
        ws = _pick_sheet(wb_value, sheet)
        sheet_name = ws.title

        if cell_range:
            rows_iter = ws[cell_range]
        else:
            rows_iter = ws.iter_rows()

        # 先取全量行（在 limit 约束下尽早停止，避免整表入内存）
        all_rows: list[list[Any]] = []
        for row in rows_iter:
            all_rows.append(_row_values(row))
            if len(all_rows) >= offset + limit + 1:
                # 多取一行用于判断是否还有剩余
                break
    finally:
        wb_value.close()

    total_rows = len(all_rows)
    page = all_rows[offset : offset + limit]
    has_more = total_rows > offset + len(page)

    # 公式读取（仅在需要时才二次打开文件，避免无谓开销）
    formulas: dict[str, str] | None = None
    if include_formula:
        formulas = {}
        wb_formula = _open(file_path, data_only=False, read_only=True)
        try:
            ws_f = _pick_sheet(wb_formula, sheet_name)
            src = ws_f[cell_range] if cell_range else ws_f.iter_rows()
            for r_i, row in enumerate(src):
                if r_i < offset:
                    continue
                if r_i >= offset + len(page):
                    break
                for c_i, cell in enumerate(row):
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formulas[cell.coordinate] = cell.value
        except Exception:
            formulas = None  # 读公式失败不阻塞主流程
        finally:
            wb_formula.close()

    # 转成 {坐标: 值} 便于 Agent 精确定位；同时给出二维行便于整体查看
    grid: list[list[Any]] = []
    cells: dict[str, Any] = {}
    start_row = offset + 1
    for r_off, row in enumerate(page, start=start_row):
        vals: list[Any] = []
        for c_off, v in enumerate(row, start=1):
            vals.append(v)
            if v is not None:
                cells[f"{get_column_letter(c_off)}{r_off}"] = v
        grid.append(vals)

    result: dict[str, Any] = {
        "file": str(file_path),
        "sheet": sheet_name,
        "range": cell_range or "(全表)",
        "offset": offset,
        "returned_rows": len(page),
        "truncated": has_more,
        "hint": (
            f"结果已截断：仅返回 {len(page)} 行（offset={offset}）。"
            f"如需更多请增大 limit 或调整 offset。"
            if has_more
            else None
        ),
        "grid": grid,
        "cells": cells,
    }
    if formulas:
        result["formulas"] = formulas
    return result


def _row_values(row) -> list[Any]:
    """取一行的值列表。

    坑：openpyxl 在任意模式下 iter_rows() / ws[range] 迭代出的都是 Cell 对象
    （read_only 模式为 ReadOnlyCell），必须取 .value 才是单元格内容，
    直接把 Cell 对象当值用会拿到 "<ReadOnlyCell '表名'.A1>" 这种字符串。
    """
    out: list[Any] = []
    for c in row:
        # 单元格对象取 value；若已是原始值（防御性）则直接用
        out.append(_clean(getattr(c, "value", c)))
    return out


def _clean(v: Any) -> Any:
    """清洗不可直接 JSON 序列化的值。"""
    if v is None:
        return None
    if isinstance(v, (str, int, float, bool)):
        # 官方 skill 提到：公式结果为 "" 时 openpyxl 也会读回 None，这里保持原样
        return v
    return str(v)


# ---------------------------------------------------------------------------
# L3 写入：bulk 批量 + 单次读写周期
# ---------------------------------------------------------------------------
def write_cells(
    file_path: Path,
    sheet: str | None,
    cells: list[dict[str, Any]],
    create_sheet_if_missing: bool = False,
) -> dict[str, Any]:
    """批量写入多个单元格，一次 load/save 完成。

    cells 形如 [{"cell": "A1", "value": "姓名"}, {"cell": "B2", "value": "=SUM(B3:B9)"}]
    - value 以 "=" 开头视为公式（学 knorq 的约定）
    - 支持跨坐标一次性写入，避免 Agent 循环调用单格接口
    """
    if not cells:
        raise ValueError("cells 为空，未提供任何要写入的单元格")
    if len(cells) > MAX_WRITE_CELLS:
        raise ValueError(
            f"单次写入 {len(cells)} 个单元格超过上限 {MAX_WRITE_CELLS}，请分批写入"
        )

    wb = _open(file_path)
    try:
        if sheet is None:
            ws = wb.active
        elif sheet in wb.sheetnames:
            ws = wb[sheet]
        elif create_sheet_if_missing:
            ws = wb.create_sheet(sheet)
        else:
            raise KeyError(
                f"工作表 '{sheet}' 不存在。可用工作表：{', '.join(wb.sheetnames)}"
                "（如需新建请传 create_sheet_if_missing=true）"
            )

        written = []
        for item in cells:
            coord = str(item.get("cell", "")).strip().upper()
            if not coord:
                raise ValueError(f"单元格项缺少 cell 坐标：{item}")
            value = item.get("value")
            # 官方 skill 的坑：空字符串与 None 语义不同，这里按原样写入
            ws[coord] = value
            written.append(
                {"cell": coord, "value": value, "is_formula": _is_formula(value)}
            )

        wb.save(str(file_path))
        return {
            "file": str(file_path),
            "sheet": ws.title,
            "written_count": len(written),
            "cells": written,
            "note": (
                "注意：公式已写入但尚未重算，读取时可能得到 None；"
                "重算能力将在 M3 阶段提供（recalculate 工具）。"
                if any(c["is_formula"] for c in written)
                else None
            ),
        }
    finally:
        wb.close()


def _is_formula(value: Any) -> bool:
    """判断值是否为 Excel 公式。"""
    return isinstance(value, str) and value.startswith("=")
