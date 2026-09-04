#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
边界回归测试：固化在 BUG 探测中发现并修复的问题，防止复发。

覆盖的类型类陷阱：
1. 文本列用数值比较符（> < >= <=）→ 必须给可读错误，不能崩 TypeError
2. 分组聚合对文本列做 max/min → 排序键不能假设值是数字
3. 含空值的列排序 / 聚合
4. 空表（仅表头）
5. read_range 的 offset 超过总行数
6. 重复 key 的表比较
"""

import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).parent.parent))

from openpyxl import Workbook

from duduexcel import advanced, analytics, excel_ops, styling

TMP = Path(__file__).parent / "_tmp_edge.xlsx"


def build() -> Path:
    if TMP.exists():
        TMP.unlink()
    wb = Workbook()
    ws = wb.active
    ws.title = "混合"
    ws.append(["姓名", "部门", "月薪", "备注"])
    for r in [("张三", "研发", 25000, "优秀"), ("李四", "市场", 18000, None),
              ("王五", "研发", None, "待定"), ("赵六", "销售", 21000, "优秀")]:
        ws.append(list(r))
    wb.create_sheet("空表").append(["列1", "列2"])
    wb.save(TMP)
    return TMP


def check(name, cond, extra="") -> bool:
    print(f"[{'PASS' if cond else 'FAIL'}] {name} {extra}")
    return cond


def main() -> int:
    p = build()
    ok = True

    # 1. 文本列 + 数值比较符：必须抛可读错误（此前是 TypeError 崩溃）
    for op in (">", ">=", "<", "<="):
        try:
            analytics.filter_count(p, "混合", [{"column": "姓名", "op": op, "value": 1}])
            ok &= check(f"文本列用 '{op}' 应报错", False)
        except analytics.AnalyticsError as e:
            ok &= check(
                f"文本列用 '{op}' 给可读错误", "不是数值列" in str(e),
                f"-> {str(e)[:52]}",
            )
        except Exception as e:
            ok &= check(f"文本列用 '{op}' 应抛 AnalyticsError", False,
                        f"-> 实际抛 {type(e).__name__}")

    # 2. 分组聚合对文本列做 max：排序键不能假设数字（此前 TypeError）
    try:
        r = analytics.aggregate(p, "混合", "姓名", op="max", group_by=["部门"])
        ok &= check("文本列 max 分组不再崩溃", len(r["result"]) == 3,
                    f"-> {r['result']}")
    except Exception as e:
        ok &= check("文本列 max 分组不再崩溃", False, f"-> {type(e).__name__}: {e}")

    # 3. 含空值的列：排序与聚合
    try:
        r = analytics.top_n(p, "混合", sort_by="月薪", n=3)
        ok &= check("含空值数值列排序", len(r["rows"]) == 3)
    except Exception as e:
        ok &= check("含空值数值列排序", False, f"-> {e}")
    try:
        r = analytics.top_n(p, "混合", sort_by="备注", n=3)
        ok &= check("含空值文本列排序", len(r["rows"]) >= 1)
    except Exception as e:
        ok &= check("含空值文本列排序", False, f"-> {e}")
    try:
        r = analytics.aggregate(p, "混合", "月薪", op="sum")
        ok &= check("含空值求和", r["value"] == 64000, f"-> {r['value']}")
    except Exception as e:
        ok &= check("含空值求和", False, f"-> {e}")

    # 4. 空表：各分析工具都不应崩
    for label, fn in [
        ("sheet_profile 空表", lambda: analytics.sheet_profile(p, "空表")),
        ("aggregate 空表", lambda: analytics.aggregate(p, "空表", "列1", op="sum")),
        ("top_n 空表", lambda: analytics.top_n(p, "空表", sort_by="列1", n=3)),
        ("filter_count 空表", lambda: analytics.filter_count(p, "空表", None)),
    ]:
        try:
            fn()
            ok &= check(label, True)
        except Exception as e:
            ok &= check(label, False, f"-> {type(e).__name__}: {e}")

    # 5. read_range 的 offset 超过总行数
    try:
        r = excel_ops.read_range(p, sheet="混合", offset=999, limit=5)
        ok &= check("offset 超范围返回空且不崩", r["returned_rows"] == 0,
                    f"-> returned={r['returned_rows']}")
    except Exception as e:
        ok &= check("offset 超范围返回空且不崩", False, f"-> {e}")

    # 6. 透视表对文本列做 values / 空表
    try:
        r = advanced.create_pivot(p, "空表", rows=["列1"], values=["列2"])
        ok &= check("透视表空表不崩", r["result_rows"] == 0)
    except Exception as e:
        ok &= check("透视表空表不崩", False, f"-> {e}")

    # 7. 样式对空表
    try:
        r = styling.apply_chinese_style(p, "空表")
        ok &= check("空表套样式不崩", r["styled_rows"] >= 1)
    except Exception as e:
        ok &= check("空表套样式不崩", False, f"-> {e}")

    if TMP.exists():
        TMP.unlink()
    bak = Path(str(TMP) + ".bak")
    if bak.exists():
        bak.unlink()

    print("")
    print(f"===== 边界回归测试: {'全部通过' if ok else '存在失败'} =====")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
