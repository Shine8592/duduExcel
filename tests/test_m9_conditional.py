#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""M9：条件格式读取（写入 + 读取闭环）。"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).parent.parent))

from openpyxl import Workbook

from duduexcel import advanced

TMP = Path(__file__).parent / "_tmp_m9.xlsx"


def check(name, cond, extra="") -> bool:
    print(f"[{'PASS' if cond else 'FAIL'}] {name} {extra}")
    return cond


def build() -> Path:
    if TMP.exists():
        TMP.unlink()
    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    ws.append(["值"])
    for v in [10, 50, 90, 120]:
        ws.append([v])
    wb.create_sheet("空表").append(["列1"])
    wb.save(TMP)
    return TMP


def main() -> int:
    ok = True
    p = build()

    # 1. 空表时无规则
    r = advanced.list_conditional_formats(p, "空表")
    ok &= check("无规则时返回 0 条", r.get("count") == 0, f"-> {r.get('count')}")

    # 2. 写入多种规则
    advanced.add_conditional_format(p, "数据", "A2:A5", "greater_than", value=80)
    advanced.add_conditional_format(p, "数据", "A2:A5", "data_bar")
    advanced.add_conditional_format(p, "数据", "A2:A5", "between", value=20, value2=100)
    advanced.add_conditional_format(p, "数据", "A2:A5", "duplicate")

    # 3. 读回
    r = advanced.list_conditional_formats(p, "数据")
    ok &= check("读回 4 条规则", r.get("count") == 4, f"-> {r.get('count')}")

    rules = r.get("rules", [])
    types = {x["type"] for x in rules}
    ok &= check("包含 cellIs / dataBar / expression 三类",
                {"cellIs", "dataBar", "expression"} <= types, f"-> {types}")

    # 4. 关键细节：阈值、运算符、填充色、优先级
    gt = next((x for x in rules if x.get("operator") == "greaterThan"), None)
    ok &= check("greaterThan 阈值读回正确", gt is not None and gt["formula"] == ["80"],
                f"-> {gt.get('formula') if gt else None}")
    ok &= check("填充色读回", gt is not None and "fill_color" in gt,
                f"-> {gt.get('fill_color') if gt else None}")

    btw = next((x for x in rules if x.get("operator") == "between"), None)
    ok &= check("between 双阈值读回正确",
                btw is not None and btw["formula"] == ["20", "100"],
                f"-> {btw.get('formula') if btw else None}")

    prios = [x.get("priority") for x in rules]
    ok &= check("按优先级排序", prios == sorted([q for q in prios if q is not None]) or all(q is None for q in prios),
                f"-> {prios}")

    # 5. 重复值规则（expression + COUNTIF 公式）
    dup = next((x for x in rules if x["type"] == "expression"), None)
    ok &= check("重复值规则公式读回", dup is not None and "COUNTIF" in "".join(dup["formula"]),
                f"-> {dup.get('formula') if dup else None}")

    # 6. 不存在的 sheet 应回退到活动表而不是崩
    try:
        r2 = advanced.list_conditional_formats(p, "不存在的表")
        ok &= check("未知表回退不崩", isinstance(r2, dict), f"-> sheet={r2.get('sheet')}")
    except Exception as e:
        ok &= check("未知表回退不崩", False, f"-> {e}")

    if TMP.exists():
        TMP.unlink()
    bak = Path(str(TMP) + ".bak")
    if bak.exists():
        bak.unlink()

    print("")
    print(f"===== M9 测试: {'全部通过' if ok else '存在失败'} =====")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
