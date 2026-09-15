#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成 duduExcel 演示文件 examples/demo.xlsx。

构造一个"带作者语义"的表格，用来展示 duduExcel 的差异点：
- 已取消的条目用删除线
- 待审阅的单元格用黄底
- 隐藏一行（敏感/不相关）
- 含公式与中文表头
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

OUT = Path(__file__).parent.parent / "examples" / "demo.xlsx"
OUT.parent.mkdir(parents=True, exist_ok=True)


def build() -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "需求清单"
    ws.append(["功能", "负责人", "状态", "工时"])
    rows = [
        ("车辆调度", "张三", "已批准", 12),
        ("旧版导出", "李四", "已取消", 5),      # 将被标删除线
        ("司机点名", "王五", "待审阅", 8),      # 状态将被标黄底
        ("报表导出", "张三", "已批准", 6),
    ]
    for r in rows:
        ws.append(list(r))

    # 作者语义 1：删除线表示已取消
    for col in ("A", "B", "C"):
        ws[f"{col}3"].font = Font(strike=True)

    # 作者语义 2：黄底表示待审阅
    ws["C4"].fill = PatternFill("solid", fgColor="FFFF00")

    # 表头加粗
    for col in "ABCD":
        ws[f"{col}1"].font = Font(bold=True)

    # 隐藏第 5 行（模拟敏感/不相关数据）
    ws.append(["内部草案", "赵六", "隐藏", 99])
    ws.row_dimensions[5].hidden = True

    # 加一张销售表，供透视表/分析演示
    ws2 = wb.create_sheet("销售")
    ws2.append(["区域", "产品", "销售额", "数量"])
    for r in [
        ("华东", "笔记本", 120, 2), ("华东", "鼠标", 35, 10),
        ("华南", "笔记本", 89, 3), ("华南", "鼠标", 45, 8),
        ("华东", "显示器", 210, 1), ("华南", "显示器", 76, 4),
    ]:
        ws2.append(list(r))

    # 加合计公式（供 recalculate 演示）
    ws2.append([None, "合计", "=SUM(C2:C7)", "=SUM(D2:D7)"])

    wb.save(OUT)
    wb.close()
    return OUT


if __name__ == "__main__":
    p = build()
    print(f"[OK] 已生成演示文件: {p}  ({p.stat().st_size/1024:.1f} KB)")
