#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""M2 服务端分析测试：验证分析语义正确 + 上下文效率（不回传数据行）。"""

import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).parent.parent))

from openpyxl import Workbook

from duduexcel import analytics

TMP = Path(__file__).parent / "_tmp_m2.xlsx"

# 造 60 行含中文、多类别、数值的数据（确保聚合/过滤/排序都能验证）
DEPTS = ["研发", "市场", "销售"]
ROWS = []


def build_fixture() -> Path:
    if TMP.exists():
        TMP.unlink()
    wb = Workbook()
    ws = wb.active
    ws.title = "员工"
    ws.append(["姓名", "部门", "月薪", "工龄"])
    for i in range(60):
        dept = DEPTS[i % 3]
        salary = 15000 + (i % 7) * 3000 if dept == "研发" else (
            12000 + (i % 5) * 2500 if dept == "市场" else 10000 + (i % 9) * 2000
        )
        ROWS.append({"dept": dept, "salary": salary})
        ws.append([f"员工{i+1}", dept, salary, (i % 10) + 1])
    wb.save(TMP)
    return TMP


def check(name: str, cond: bool, extra: str = "") -> bool:
    print(f"[{'PASS' if cond else 'FAIL'}] {name} {extra}")
    return cond


def expected(dept: str, op: str):
    vals = [r["salary"] for r in ROWS if r["dept"] == dept]
    if op == "sum":
        return sum(vals)
    if op == "mean":
        return round(sum(vals) / len(vals), 4)
    if op == "count":
        return len(vals)
    if op == "max":
        return max(vals)
    if op == "min":
        return min(vals)
    if op == "nunique":
        return len(set(vals))
    return None


def main() -> int:
    p = build_fixture()
    ok = True

    # 1. sheet_profile：一次调用摸清全表
    prof = analytics.sheet_profile(p, "员工")
    ok &= check("sheet_profile 识别出 4 列", len(prof["profile"]) == 4,
                f"-> {[c['column'] for c in prof['profile']]}")
    salary_col = next(c for c in prof["profile"] if c["column"] == "月薪")
    ok &= check("月薪列有数值统计", "stats" in salary_col, f"-> stats={salary_col.get('stats')}")
    dept_col = next(c for c in prof["profile"] if c["column"] == "部门")
    ok &= check("部门列识别出 3 个类别", dept_col["unique"] == 3, f"-> unique={dept_col['unique']}")
    ok &= check("sheet_profile 自报 token 节省", prof["_meta"]["tokens_saved"] > 0,
                f"-> saved={prof['_meta']['tokens_saved']} ({prof['_meta']['savings_pct']}%)")
    # 关键：画像里不应包含原始数据行（"rows" 是行数标量，不是行列表）
    ok &= check(
        "sheet_profile 不回传数据行",
        "grid" not in prof and not isinstance(prof.get("rows"), list),
        f"-> rows={prof.get('rows')} (标量行数，非行列表)",
    )

    # 2. filter_count：只给计数
    fc = analytics.filter_count(p, "员工", [{"column": "部门", "op": "==", "value": "研发"}])
    exp_count = sum(1 for r in ROWS if r["dept"] == "研发")
    ok &= check("filter_count 计数正确", fc["matched_rows"] == exp_count,
                f"-> {fc['matched_rows']} (期望 {exp_count})")
    ok &= check("filter_count 给出百分比", 0 < fc["matched_pct"] < 100, f"-> {fc['matched_pct']}%")
    ok &= check("filter_count 只给少量样例", len(fc["sample_rows"]) <= 3,
                f"-> {len(fc['sample_rows'])} 条")
    ok &= check("filter_count 附 Excel 公式", "COUNTIFS" in (fc["excel_formula"] or ""),
                f"-> {fc['excel_formula']}")

    # 多条件 AND
    fc2 = analytics.filter_count(p, "员工", [
        {"column": "部门", "op": "==", "value": "研发"},
        {"column": "月薪", "op": ">", "value": 20000},
    ])
    exp2 = sum(1 for r in ROWS if r["dept"] == "研发" and r["salary"] > 20000)
    ok &= check("多条件 AND 计数正确", fc2["matched_rows"] == exp2,
                f"-> {fc2['matched_rows']} (期望 {exp2})")

    # 3. aggregate：不分组
    agg = analytics.aggregate(p, "员工", "月薪", op="sum")
    exp_sum = sum(r["salary"] for r in ROWS)
    ok &= check("aggregate sum 正确", agg["value"] == exp_sum, f"-> {agg['value']} (期望 {exp_sum})")
    ok &= check("aggregate 附 Excel 公式", "SUM" in (agg["excel_formula"] or ""))

    # 4. aggregate：分组（透视）
    aggg = analytics.aggregate(p, "员工", "月薪", op="sum", group_by=["部门"])
    got = {r["部门"]: r["sum"] for r in aggg["result"]}
    ok &= check("分组聚合组数正确", len(aggg["result"]) == 3, f"-> {list(got.keys())}")
    ok &= check("分组 sum 各值正确",
                all(got.get(d) == expected(d, "sum") for d in DEPTS),
                f"-> got={got}")
    ok &= check("分组结果附 tsv", bool(aggg["tsv"]) and "\t" in aggg["tsv"])
    ok &= check("分组结果只回传每组一个数字（非全表行）",
                len(aggg["result"]) == 3 and aggg["rows_scanned"] == 60,
                f"-> groups={len(aggg['result'])} scanned={aggg['rows_scanned']}")

    # 5. top_n：排行榜
    tn = analytics.top_n(p, "员工", sort_by="月薪", n=5)
    ok &= check("top_n 返回 5 行", len(tn["rows"]) == 5)
    ok &= check("top_n 降序排列",
                all(tn["rows"][i]["月薪"] >= tn["rows"][i + 1]["月薪"] for i in range(4)),
                f"-> {[r['月薪'] for r in tn['rows']]}")
    ok &= check("top_n 含 rank 字段", tn["rows"][0]["rank"] == 1)
    ok &= check("top_n 只回传 5 行而非全表 60 行", tn["rows_scanned"] == 60 and tn["returned"] == 5)

    # 升序取最小 + 只返回指定列（省 token）
    tn2 = analytics.top_n(p, "员工", sort_by="月薪", n=3, ascending=True, columns=["姓名", "月薪"])
    ok &= check("升序取最小值正确", tn2["rows"][0]["月薪"] == min(r["salary"] for r in ROWS),
                f"-> {[r['月薪'] for r in tn2['rows']]}")
    ok &= check("columns 参数只回传指定列", set(tn2["rows"][0].keys()) == {"rank", "姓名", "月薪"},
                f"-> {list(tn2['rows'][0].keys())}")

    # 6. 错误处理
    try:
        analytics.aggregate(p, "员工", "不存在的列", op="sum")
        ok &= check("错误列名应报错", False)
    except Exception as e:
        ok &= check("错误列名给出可用列", "可用列" in str(e), f"-> {str(e)[:60]}")

    try:
        analytics.filter_count(p, "员工", [{"column": "部门", "op": "~~", "value": "x"}])
        ok &= check("错误运算符应报错", False)
    except Exception as e:
        ok &= check("错误运算符给出可用清单", "不支持的运算符" in str(e), f"-> {str(e)[:60]}")

    if TMP.exists():
        TMP.unlink()

    print("")
    print(f"===== M2 分析测试: {'全部通过' if ok else '存在失败'} =====")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
