#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""M1 冒烟测试：构造真实 xlsx → 直接调用核心函数验证三个工具。"""

import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).parent.parent))

from openpyxl import Workbook

from duduexcel import excel_ops
from duduexcel.safety import backup_file, resolve_path, revert_last

TMP = Path(__file__).parent / "_tmp_smoke.xlsx"


def build_fixture() -> Path:
    """构造含中文表头、数值、公式、多工作表的测试文件。"""
    if TMP.exists():
        TMP.unlink()
    wb = Workbook()
    ws = wb.active
    ws.title = "销售明细"
    ws.append(["产品", "单价", "数量", "金额"])
    rows = [
        ["笔记本", 5999, 3, "=B2*C2"],
        ["鼠标", 129, 20, "=B3*C3"],
        ["显示器", 1899, 5, "=B4*C4"],
        ["键盘", 349, 12, "=B5*C5"],
    ]
    for r in rows:
        ws.append(r)
    ws2 = wb.create_sheet("汇总")
    ws2["A1"] = "总计"
    ws2["B1"] = "=SUM(销售明细!D2:D5)"
    wb.save(TMP)
    return TMP


def check(name: str, cond: bool, extra: str = "") -> bool:
    print(f"[{'PASS' if cond else 'FAIL'}] {name} {extra}")
    return cond


def main() -> int:
    p = build_fixture()
    ok = True

    # 1. workbook_info
    info = excel_ops.inspect_workbook(p)
    ok &= check(
        "workbook_info 返回两个工作表",
        info["sheet_count"] == 2 and "销售明细" in [s["name"] for s in info["sheets"]],
        f"-> sheets={[s['name'] for s in info['sheets']]}",
    )
    ok &= check(
        "workbook_info 行列数正确",
        info["sheets"][0]["rows"] == 5 and info["sheets"][0]["columns"] == 4,
        f"-> {info['sheets'][0]['rows']}行 x {info['sheets'][0]['columns']}列",
    )

    # 2. read_range：默认 limit 截断能力
    r = excel_ops.read_range(p, sheet="销售明细", limit=3)
    ok &= check("read_range limit=3 只返回 3 行", len(r["grid"]) == 3)
    ok &= check("read_range 截断标志正确", r["truncated"] is True, f"-> truncated={r['truncated']}")
    ok &= check("read_range 提示信息存在", bool(r["hint"]))
    ok &= check(
        "read_range 读到中文表头",
        r["grid"][0][0] == "产品",
        f"-> 首行={r['grid'][0]}",
    )
    ok &= check(
        "read_range 同时给出公式",
        "D2" in (r.get("formulas") or {}) or "=B2*C2" in str(r.get("formulas")),
        f"-> formulas={r.get('formulas')}",
    )

    # 分页：共 5 行（1 表头 + 4 数据），offset=3 跳过表头与"笔记本""鼠标"，落到"显示器"
    r2 = excel_ops.read_range(p, sheet="销售明细", offset=3, limit=3)
    ok &= check(
        "read_range offset 分页生效",
        r2["grid"][0][0] == "显示器" and r2["truncated"] is False,
        f"-> 首行={r2['grid'][0]}, returned={r2['returned_rows']}",
    )

    # 公式未重算时读到 None 是预期行为（M3 提供 recalculate 后才有缓存值）
    ok &= check(
        "未重算的公式读到 None（预期行为）",
        r2["grid"][0][3] is None,
        f"-> 金额列={r2['grid'][0][3]}",
    )

    # 3. write_cells：批量写入 + 备份
    bak = backup_file(p)
    res = excel_ops.write_cells(
        p,
        sheet="销售明细",
        cells=[
            {"cell": "A6", "value": "耳机"},
            {"cell": "B6", "value": 899},
            {"cell": "C6", "value": 8},
            {"cell": "D6", "value": "=B6*C6"},
        ],
    )
    ok &= check("write_cells 批量写入 4 格", res["written_count"] == 4)
    ok &= check("write_cells 识别公式", res["cells"][3]["is_formula"] is True)
    ok &= check("write_cells 提示未重算", bool(res["note"]))

    # 验证落盘
    r3 = excel_ops.read_range(p, sheet="销售明细", offset=5, limit=2)
    ok &= check(
        "写入内容已落盘",
        r3["grid"][0][0] == "耳机" and r3["grid"][0][1] == 899,
        f"-> {r3['grid'][0] if r3['grid'] else '空'}",
    )

    # 4. 回滚
    revert_last(p)
    r4 = excel_ops.read_range(p, sheet="销售明细", limit=10)
    ok &= check(
        "revert_last 回滚成功（耳机已消失）",
        all(row[0] != "耳机" for row in r4["grid"] if row),
    )

    # 5. 错误处理：不存在的 sheet 给出可读提示
    try:
        excel_ops.read_range(p, sheet="不存在的表", limit=2)
        ok &= check("错误 sheet 应抛异常", False)
    except KeyError as e:
        ok &= check(
            "错误 sheet 名给出可用列表", "销售明细" in str(e), f"-> {str(e)[:60]}"
        )

    # 6. 安全层：路径白名单
    import os
    os.environ["DUDU_EXCEL_ROOT"] = str(Path(__file__).parent)
    try:
        resolve_path("../../Windows/System32/config/SAM")
        ok &= check("目录穿越应被拒绝", False)
    except Exception as e:
        ok &= check("目录穿越被拒绝", "越界" in str(e), f"-> {str(e)[:60]}")
    finally:
        os.environ.pop("DUDU_EXCEL_ROOT", None)

    # 清理
    if TMP.exists():
        TMP.unlink()
    bak_p = Path(bak)
    if bak_p.exists():
        bak_p.unlink()

    print("")
    print(f"===== M1 冒烟测试: {'全部通过' if ok else '存在失败'} =====")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
