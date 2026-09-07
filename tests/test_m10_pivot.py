#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
M10：可交互透视表（真 PivotTable）。

验证策略（不依赖 openpyxl 的保守解析）：
1. zip 内必须存在 5 个 pivot 部件
2. 目标 sheet 的 rels 必须引用 pivotTable（决定 Excel 能否识别）
3. workbook.xml 必须注册 <pivotCaches>
4. [Content_Types].xml 必须声明三个 Override
5. 用 LibreOffice 真实打开重存后 pivot 部件仍在（最强验证，需装 LibreOffice）
"""
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).parent.parent))

from openpyxl import Workbook

from duduexcel import pivot_ooxml

TMP = Path(__file__).parent / "_tmp_m10.xlsx"
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"


def check(name, cond, extra="") -> bool:
    print(f"[{'PASS' if cond else 'FAIL'}] {name} {extra}")
    return cond


def build(target: Path | None = None) -> Path:
    """生成测试用源文件；不传 target 时用默认 TMP。"""
    out = target or TMP
    if out.exists():
        out.unlink()
    wb = Workbook()
    ws = wb.active
    ws.title = "销售"
    ws.append(["区域", "产品", "销售额", "数量"])
    for r in [("华东", "笔记本", 120, 2), ("华东", "鼠标", 35, 10),
              ("华南", "笔记本", 89, 3), ("华南", "鼠标", 45, 8),
              ("华东", "显示器", 210, 1), ("华南", "显示器", 76, 4)]:
        ws.append(list(r))
    wb.save(out)
    wb.close()          # 显式关闭，避免 Windows 下文件句柄占用导致 unlink 失败
    return out


def main() -> int:
    ok = True
    p = build()

    # ---------- 1. 基础：行 + 值 ----------
    res = pivot_ooxml.create_interactive_pivot(
        p, source_sheet="销售", rows=["区域"], values=["销售额"], agg_func="sum",
        target_sheet="PT基础",
    )
    ok &= check("基础透视表构造成功", res.get("rows_aggregated") == 6,
                f"-> 聚合 {res.get('rows_aggregated')} 行")

    # ---------- 2. zip 部件齐全 ----------
    with zipfile.ZipFile(p) as z:
        names = z.namelist()
        required = [
            "xl/pivotTables/pivotTable1.xml",
            "xl/pivotCache/pivotCacheDefinition1.xml",
            "xl/pivotCache/pivotCacheRecords1.xml",
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
        ]
        missing = [r for r in required if r not in names]
        ok &= check("pivot 部件齐全", not missing, f"-> 缺失 {missing or '无'}")

        pt_xml = z.read("xl/pivotTables/pivotTable1.xml").decode("utf-8")
        ok &= check("pivotTable 含 rowFields", "<rowFields" in pt_xml)
        ok &= check("pivotTable 含 dataFields", "<dataFields" in pt_xml)
        ok &= check("sum 聚合映射正确", 'subtotal="sum"' in pt_xml)

        ct = z.read("[Content_Types].xml").decode("utf-8")
        ok &= check("Content_Types 声明 3 个 Override",
                    ct.count("pivotCacheDefinition+xml") == 1
                    and ct.count("pivotCacheRecords+xml") == 1
                    and ct.count("pivotTable+xml") == 1)
        wbx = z.read("xl/workbook.xml").decode("utf-8")
        ok &= check("workbook.xml 注册 pivotCaches", "<pivotCaches>" in wbx)
        wb_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        ok &= check("workbook.rels 指向 pivotCache",
                    "pivotCacheDefinition" in wb_rels)

        # 关键：sheet rels 必须挂上 pivotTable
        rels_hit = [
            n for n in names
            if n.startswith("xl/worksheets/_rels/")
            and "pivotTable" in z.read(n).decode("utf-8", "replace")
        ]
        ok &= check("目标 sheet rels 引用 pivotTable（决定 Excel 能否识别）",
                    len(rels_hit) > 0, f"-> {rels_hit}")

    # ---------- 3. 交叉表（列字段） ----------
    p2 = Path(str(TMP).replace("_tmp_m10", "_tmp_m10b"))
    build(p2)
    res2 = pivot_ooxml.create_interactive_pivot(
        p2, source_sheet="销售", rows=["区域"], columns=["产品"],
        values=["销售额"], agg_func="sum", target_sheet="PT交叉",
    )
    with zipfile.ZipFile(p2) as z:
        pt2 = z.read("xl/pivotTables/pivotTable1.xml").decode("utf-8")
    ok &= check("交叉表含 colFields", "<colFields" in pt2,
                f"-> rows_aggregated={res2.get('rows_aggregated')}")
    try:
        p2.unlink()
    except Exception:
        pass

    # ---------- 4. 筛选字段 + count ----------
    p3 = Path(str(TMP).replace("_tmp_m10", "_tmp_m10c"))
    build(p3)
    pivot_ooxml.create_interactive_pivot(
        p3, source_sheet="销售", rows=["区域"], values=["数量"],
        agg_func="count", page_fields=["产品"], target_sheet="PT筛选",
    )
    with zipfile.ZipFile(p3) as z:
        pt3 = z.read("xl/pivotTables/pivotTable1.xml").decode("utf-8")
    ok &= check("筛选字段含 pageFields", "<pageFields" in pt3)
    ok &= check("count 聚合映射正确", 'subtotal="count"' in pt3)
    try:
        p3.unlink()
    except Exception:
        pass

    # ---------- 5. 行过滤条件 ----------
    p4 = Path(str(TMP).replace("_tmp_m10", "_tmp_m10d"))
    build(p4)
    res4 = pivot_ooxml.create_interactive_pivot(
        p4, source_sheet="销售", rows=["区域"], values=["销售额"],
        filters=[{"column": "区域", "op": "==", "value": "华东"}],
        target_sheet="PT过滤",
    )
    ok &= check("行过滤生效（华东 3 行）", res4.get("rows_aggregated") == 3,
                f"-> {res4.get('rows_aggregated')}")
    try:
        p4.unlink()
    except Exception:
        pass

    # ---------- 6. 错误参数 ----------
    p5 = Path(str(TMP).replace("_tmp_m10", "_tmp_m10e"))
    build(p5)
    try:
        pivot_ooxml.create_interactive_pivot(p5, values=["不存在列"], rows=["区域"])
        ok &= check("不存在的字段应报错", False)
    except pivot_ooxml.PivotError as e:
        ok &= check("不存在的字段给出可读错误", "字段不存在" in str(e), f"-> {str(e)[:50]}")
    try:
        pivot_ooxml.create_interactive_pivot(p5, values=["销售额"], rows=["区域"], agg_func="中位数")
        ok &= check("不支持的聚合应报错", False)
    except pivot_ooxml.PivotError as e:
        ok &= check("不支持的聚合给出可用清单", "不支持的聚合函数" in str(e))
    try:
        p5.unlink()
    except Exception:
        pass

    # ---------- 7. LibreOffice 真实打开验证（最强证据） ----------
    if Path(SOFFICE).exists():
        outdir = Path(__file__).parent / "_lo_out"
        if outdir.exists():
            shutil.rmtree(outdir)
        outdir.mkdir()
        try:
            r = subprocess.run(
                [SOFFICE, "--headless", "--norestore", "--convert-to", "xlsx",
                 "--outdir", str(outdir), str(p)],
                capture_output=True, text=True, timeout=240,
            )
            saved = outdir / p.name
            if saved.exists():
                with zipfile.ZipFile(saved) as z:
                    kept = [n for n in z.namelist() if "pivot" in n.lower()]
                ok &= check("LibreOffice 打开后仍保留 pivot 部件（真实 OOXML 验证）",
                            len(kept) >= 3, f"-> 保留 {len(kept)} 个")
            else:
                ok &= check("LibreOffice 转换输出存在", False)
        except Exception as e:
            ok &= check("LibreOffice 验证", False, f"-> {e}")
        finally:
            if outdir.exists():
                shutil.rmtree(outdir, ignore_errors=True)
    else:
        print("[SKIP] 未安装 LibreOffice，跳过真实打开验证")

    if TMP.exists():
        try:
            TMP.unlink()
        except Exception:
            pass

    print()
    print(f"===== M10 测试: {'全部通过' if ok else '存在失败'} =====")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
