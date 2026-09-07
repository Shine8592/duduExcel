#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
duduExcel MCP 服务入口。

设计上贯彻两条主线：
- 上下文高效：分页/截断/服务端聚合，绝不把整表灌进 Agent 上下文
- 安全：路径沙箱 + 写前备份 + 原子保存 + 失败回滚

工具描述（docstring）是 Agent 唯一能看到的说明书，因此每条都写明：
用途 / 何时用 / 批量优先提醒 / 限制与截断行为。
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Optional

# 启动早期把脚本目录加入 sys.path，兼容以 stdio 被宿主拉起、
# 不继承 PYTHONPATH 的场景（沿用记忆系统 mcp_server.py 的同款处理）
sys.path.insert(0, str(Path(__file__).parent))

try:
    from mcp.server.mcpserver import MCPServer
except ImportError as e:  # pragma: no cover
    raise SystemExit(
        f"缺少 mcp 依赖：{e}\n请执行：pip install 'mcp>=2.0.0' openpyxl"
    )

from duduexcel import advanced, analytics, excel_ops, pivot_ooxml, recalc, styling
from duduexcel.safety import (
    SafetyError,
    backup_file,
    require_exists,
    resolve_path,
    revert_last,
)

SERVER_VERSION = "0.2.0"

mcp = MCPServer("duduExcel")


# 注意：这是内部辅助函数，**不要**加 @mcp.tool() 装饰器
# （此前误加导致它被注册成一个名为 _as_bool 的工具，污染了 tools/list）
def _as_bool(v, default: bool) -> bool:
    """把参数稳健地转成布尔值。

    坑：MCP SDK（mcp 2.x）对带默认值的 bool 参数，在 schema→函数调用绑定时
    可能不传递该键，导致工具内始终拿到默认值（表现为"传了 True 却按 False 跑"）。
    因此这里同时接受 bool、字符串（"true"/"1"/"yes"）与 None。
    """
    if v is None:
        return default
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    if isinstance(v, str):
        s = v.strip().lower()
        if s in ("true", "1", "yes", "y", "on"):
            return True
        if s in ("false", "0", "no", "n", "off"):
            return False
    return default


@mcp.tool()
def workbook_info(file_path: str) -> dict:
    """查看 Excel 工作簿结构：工作表清单、每张表的行列数、合并单元格数、隐藏行列数、内嵌图片数与文件大小。

    用途：开始处理任何 Excel 文件前**先调用这个**，用一次调用摸清结构，
    避免为了知道有哪些表就把整张表读进上下文。

    返回：文件路径、大小、活动表名、每个工作表的 name/rows/columns/merged_cells/hidden_rows/hidden_columns。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return excel_ops.inspect_workbook(p)


@mcp.tool()
def write_cells(
    file_path: str,
    cells: list[dict],
    sheet: Optional[str] = None,
    create_sheet_if_missing: bool = False,
) -> dict:
    """批量写入多个单元格 —— 一次调用完成，不要用单格写入循环调用。

    这是**批量接口**：请在一次调用里传入所有要写的单元格，
    服务只做一次文件读写周期。逐格循环调用会慢很多且反复触发备份。

    参数：
    - cells：列表，每项 {"cell": "B2", "value": "内容"}；
      value 以 "=" 开头会被当作公式写入（如 "=SUM(B3:B9)"）
    - sheet：工作表名，省略则写活动表
    - create_sheet_if_missing：工作表不存在时是否新建（默认 False，避免误建表）

    安全：写入前自动生成 .bak 备份 + 原子保存，可用 revert_last_write 回滚。
    限制：单次最多 10 万个单元格。
    """
    p = resolve_path(file_path)
    require_exists(p)
    backup = backup_file(p)  # 写操作不可逆，先落备份
    try:
        return excel_ops.write_cells(
            p,
            sheet=sheet,
            cells=cells,
            create_sheet_if_missing=create_sheet_if_missing,
        )
    except Exception:
        # 写入失败：自动回滚，不让文件停留在半成品状态
        revert_last(p)
        raise


@mcp.tool()
def read_range(
    file_path: str,
    sheet: Optional[str] = None,
    cell_range: Optional[str] = None,
    offset: int = 0,
    limit: int = 200,
    include_format: bool = True,
    include_hidden: bool = False,
) -> dict:
    """读取工作表数据，支持分页、截断、格式语义与隐藏内容处理。

    重要：不要试图一次读完大表。默认 limit=200 行，
    超出的部分通过增大 offset 分批获取（truncated 会告诉你是否还有剩余）。

    参数：
    - sheet：工作表名，省略则用活动表
    - cell_range：Excel 区域，如 "A1:D50"；省略则读整表（仍受 limit 约束）
    - offset / limit：分页，从第 offset 行开始取 limit 行
    - include_format：是否附加格式语义标记（默认 True）
    - include_hidden：是否读取隐藏行列（默认 False）

    格式语义标记（汲取自 excel-vision-mcp —— 作者留下的格式是有含义的）：
    - [B] 粗体 / [I] 斜体 / [S] **删除线（常表示"已取消/作废"）**
    - [HL:色] 底色高亮（常表示"待审阅/重点"） / [C:色] 字体色
    - [M] 合并单元格 / [HIDDEN-REF] 已隐藏但因被公式引用而保留
    注意：只在单元格真的用了格式时才附加，朴素表格不产生任何额外 token。

    隐藏内容：默认跳过（作者隐藏通常表示不属于审阅内容），
    但**被可见公式引用的隐藏单元格仍保留**并标记 [HIDDEN-REF]，
    且始终报告跳过了多少，绝不静默丢弃。

    返回：grid（二维值）、cells（坐标→值）、formulas、format_markers。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return excel_ops.read_range(
        p, sheet=sheet, cell_range=cell_range, offset=offset, limit=limit,
        include_format=_as_bool(include_format, True),
        include_hidden=_as_bool(include_hidden, False),
    )


@mcp.tool()
def list_images(file_path: str) -> dict:
    """列出工作簿内的嵌图片（零依赖，直接扫描 xl/media/）。

    为什么需要它：多数 Excel MCP 会**静默丢弃所有内嵌图片**，
    于是贴在单元格里的流程图/截图永远到不了模型眼前。

    返回：图片数量、文件名、大小、尺寸（PNG/JPEG/GIF/BMP 可从文件头解析）。
    注意：只返回**清单**而非图片本体，避免 base64 撑爆上下文；
    若需模型"看到"图片内容，请解压后交给多模态模型读取。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return excel_ops.list_images(p)


@mcp.tool()
def sheet_profile(file_path: str, sheet: Optional[str] = None) -> dict:
    """查看工作表的列级画像：类型、空值率、唯一数、Top 值、数值统计。

    **这是开始分析任何表格前最该先调用的工具** —— 一次调用即可摸清全表
    （列类型、哪些列有空值、数值范围、类别分布），
    省掉反复试探的十几次调用，且不会把任何数据行塞进上下文。

    参数：
    - sheet：工作表名，省略则用活动表（首行作为表头）

    返回：每列的 profile（dtype/null_pct/unique/top_values/stats），
    以及 _meta.tokens_saved（本次相比"读全表"节省了多少 token）。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return analytics.sheet_profile(p, sheet)


@mcp.tool()
def filter_count(
    file_path: str,
    filters: Optional[list[dict]] = None,
    sheet: Optional[str] = None,
    sample: int = 3,
) -> dict:
    """按条件统计行数 —— 只返回计数，不返回数据行。

    用途：回答"有多少条满足 X"这类问题。计数在服务端完成，
    绝不会把成千上万行搬进上下文。

    参数：
    - filters：条件列表，每项 {"column": "部门", "op": "==", "value": "研发"}
      支持运算符：== != > < >= <= in not_in contains startswith endswith is_null not_null
    - sample：附带返回的样例行数（默认 3，仅用于确认过滤是否符合预期）

    返回：matched_rows、matched_pct、少量样例，以及等价的 Excel COUNTIFS 公式
    （可粘回表格动态更新）。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return analytics.filter_count(p, sheet, filters, sample=sample)


@mcp.tool()
def aggregate(
    file_path: str,
    column: str,
    op: str = "sum",
    group_by: Optional[list[str]] = None,
    filters: Optional[list[dict]] = None,
    sheet: Optional[str] = None,
    top_n: Optional[int] = None,
) -> dict:
    """服务端聚合：sum / mean / median / min / max / count / std / var / nunique。

    关键：聚合在服务端算好，只回传"每组一个数字"。
    不要让模型自己遍历数据求和 —— 大数相加模型极易算错。

    参数：
    - column：要聚合的列
    - op：聚合运算（默认 sum）
    - group_by：分组列（可多列），给出透视表式结果
    - filters：过滤条件（格式同 filter_count）
    - top_n：分组结果只返回前 N 组（按聚合值降序）

    返回：聚合值或分组表，附 tsv（可直接粘回 Excel）与等价 Excel 公式。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return analytics.aggregate(
        p, sheet, column, op=op, group_by=group_by, filters=filters, top_n=top_n
    )


@mcp.tool()
def top_n(
    file_path: str,
    sort_by: str,
    n: int = 10,
    ascending: bool = False,
    columns: Optional[list[str]] = None,
    filters: Optional[list[dict]] = None,
    sheet: Optional[str] = None,
) -> dict:
    """取排序后的前 N 行（默认降序），只回传这 N 行。

    用途：排行榜类问题（"销售额前 10 的产品"）。
    排序在服务端完成，只把 TopN 行回传，而不是把全表交给模型排序。

    参数：
    - sort_by：排序列
    - n：返回行数（默认 10）
    - ascending：True 为升序（取最小的），默认 False 降序
    - columns：只返回这些列（省 token），省略返回全部列
    - filters：过滤条件（格式同 filter_count）

    返回：带 rank 的行列表、tsv、以及等价的 Excel RANK 公式。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return analytics.top_n(
        p, sheet, sort_by, n=n, ascending=ascending, columns=columns, filters=filters
    )


@mcp.tool()
def recalculate(file_path: str, timeout: int = 60, force: bool = False) -> dict:
    """重算公式并扫描错误 —— **交付前必做的验证步骤**（差异化能力，多数同类项目不支持）。

    为什么需要它：openpyxl 写入的公式**不带计算结果**，
    直接用 read_range 读会返回 None。重算能填上缓存值并暴露公式错误。

    安全护栏（汲取官方 xlsx skill 的 recalc 思想）：
    - **外链熔断**：若工作簿引用了外部文件且缓存值已丢失，
      重算会把它们变成 #NAME? 并**永久删除外链**（不可逆）。
      此时本工具默认**拒绝执行**并列出风险项，确认接受损失才传 force=true。
    - **静默失败防护**：比对文件指纹，若 LibreOffice 正常退出却没重写文件会明确报错。

    参数：
    - timeout：超时秒数（默认 60，大工作簿请加大）
    - force：忽略外链风险强制执行（默认 False）

    返回：total_errors、errors_by_type（含位置）、以及"重算通过≠结果正确"的提醒。
    环境未装 LibreOffice 时会明确告知，不会静默假装成功。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return recalc.recalculate(p, timeout=timeout, force=force)


@mcp.tool()
def scan_formula_errors(file_path: str) -> dict:
    """只扫描公式错误（#VALUE! / #DIV/0! / #REF! / #NAME? / #NULL! / #NUM! / #N/A），不重算。

    用途：快速体检现有文件是否有坏公式。
    返回 total_errors 与按类型分组的位置明细；位置列表超过 100 条时会
    明确说明被截断数量（请以 total_errors 为准，不要以列表长度判断严重程度）。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return recalc.scan_formula_errors(p)


@mcp.tool()
def apply_chinese_style(
    file_path: str,
    sheet: Optional[str] = None,
    header_row: int = 1,
    freeze_header: bool = True,
    auto_width: bool = True,
    font_name: Optional[str] = None,
) -> dict:
    """套用中文场景样式：中文字体表头、自动列宽（中文按 2 字符宽）、冻结首行、细边框。

    为什么需要它：通用规范常推荐 Arial/Times New Roman，但对中文表格并不适用。
    本工具默认使用**微软雅黑**（屏幕阅读友好），并让中文列宽自适应（中文按 2 字符宽计算）。

    参数：
    - header_row：表头行号（默认 1）
    - auto_width：是否自适应列宽（默认 True）
    - freeze_header：是否冻结表头（默认 True）
    - font_name：自定义字体，省略用微软雅黑
    """
    p = resolve_path(file_path)
    require_exists(p)
    backup_file(p)
    try:
        return styling.apply_chinese_style(
            p, sheet, header_row=header_row, freeze_header=freeze_header,
            auto_width=auto_width, font_name=font_name,
        )
    except Exception:
        revert_last(p)
        raise


@mcp.tool()
def set_number_format(
    file_path: str,
    cell_range: str,
    number_format: str,
    sheet: Optional[str] = None,
) -> dict:
    """设置数字格式。

    number_format 可传内置名或自定义格式码：
    - cny → ¥#,##0（人民币整数）
    - cny2 → ¥#,##0.00
    - cny_dash → ¥#,##0;(¥#,##0);-（零显示为短横）
    - percent → 0.0%（注意：值必须存小数，存 0.15 显示 15.0%）
    - multiple → 0.0x（估值倍数）
    - int → #,##0；date → yyyy-mm-dd
    """
    p = resolve_path(file_path)
    require_exists(p)
    backup_file(p)
    try:
        return styling.set_number_format(p, sheet, cell_range, number_format)
    except Exception:
        revert_last(p)
        raise


@mcp.tool()
def add_chart(
    file_path: str,
    data_range: str,
    sheet: Optional[str] = None,
    chart_type: str = "bar",
    categories_range: Optional[str] = None,
    title: str = "",
    anchor_cell: str = "E2",
) -> dict:
    """插入图表（差异化能力：官方 xlsx skill 无图表指导，竞品 knorq 明确不支持图表）。

    参数：
    - data_range：数据区域，如 "B2:B10"
    - chart_type：bar / line / pie / scatter（默认 bar）
    - categories_range：分类轴区域，如 "A2:A10"
    - title：图表标题
    - anchor_cell：图表左上角锚定位置（默认 E2）
    """
    p = resolve_path(file_path)
    require_exists(p)
    backup_file(p)
    try:
        return styling.add_chart(
            p, sheet, chart_type=chart_type, data_range=data_range,
            categories_range=categories_range, title=title, anchor_cell=anchor_cell,
        )
    except Exception:
        revert_last(p)
        raise


@mcp.tool()
def create_pivot(
    file_path: str,
    rows: list[str],
    values: list[str],
    agg_func: str = "sum",
    columns: Optional[list[str]] = None,
    filters: Optional[list[dict]] = None,
    source_sheet: Optional[str] = None,
    target_sheet: str = "透视表",
) -> dict:
    """生成透视汇总表（写入新工作表）。

    ⚠️ 诚实说明：openpyxl 无法创建真正可交互的 PivotTable 对象，
    本工具生成的是**静态汇总表**（分组聚合后写回），数值等价但不可交互。
    需要可交互透视表时，请在 Excel 中基于结果表插入。

    参数：
    - rows：行分组字段（可多列），如 ["部门"]
    - values：要聚合的数值列，如 ["销售额"]
    - agg_func：sum/mean/count/min/max（默认 sum）
    - columns：列分组字段（可选，做交叉表）
    - filters：过滤条件（格式同 filter_count）
    - target_sheet：结果表名（默认"透视表"）
    """
    p = resolve_path(file_path)
    require_exists(p)
    backup_file(p)
    try:
        return advanced.create_pivot(
            p, source_sheet, rows, values, agg_func=agg_func,
            columns=columns, filters=filters, target_sheet=target_sheet,
        )
    except Exception:
        revert_last(p)
        raise


@mcp.tool()
def add_conditional_format(
    file_path: str,
    cell_range: str,
    cond_type: str,
    value: Optional[float] = None,
    value2: Optional[float] = None,
    sheet: Optional[str] = None,
) -> dict:
    """添加条件格式（openpyxl 原生规则，真实生效）。

    cond_type 可选：
    - data_bar：数据条（无需 value）
    - color_scale：色阶（无需 value）
    - greater_than / less_than / equal：需要 value
    - between：需要 value 与 value2
    - duplicate：高亮重复值（无需 value）

    参数：
    - cell_range：作用区域，如 "C2:C100"
    - value / value2：阈值
    """
    p = resolve_path(file_path)
    require_exists(p)
    backup_file(p)
    try:
        return advanced.add_conditional_format(
            p, sheet, cell_range, cond_type, value=value, value2=value2
        )
    except Exception:
        revert_last(p)
        raise


@mcp.tool()
def create_interactive_pivot(
    file_path: str,
    rows: list[str],
    values: list[str],
    agg_func: str = "sum",
    columns: Optional[list[str]] = None,
    page_fields: Optional[list[str]] = None,
    filters: Optional[list[dict]] = None,
    source_sheet: Optional[str] = None,
    target_sheet: str = "PivotTable",
    location: str = "A3",
) -> dict:
    """Create an INTERACTIVE PivotTable (a real Excel PivotTable object).

    Difference from `create_pivot` (static summary):
    - `create_pivot`           : writes aggregated numbers into cells - fast but static
    - `create_interactive_pivot`: injects real OOXML pivot parts, so in Excel you can
      drag fields, expand/collapse and refresh (structure validated with LibreOffice)

    Args:
    - rows / columns / values: row / column (cross-tab) / value fields
    - agg_func: sum / count / average / min / max (default sum)
    - page_fields: report filter fields
    - filters: row filters (same format as filter_count)
    - target_sheet: target worksheet (created if missing)
    - location: top-left anchor (default A3)

    LIMITATIONS:
    1. Saving this file again with openpyxl (incl. this server's write_cells) will DROP
       the PivotTable - openpyxl cannot write pivot parts back. Generate it LAST.
    2. Not supported: field grouping, calculated fields/items, slicers, timelines,
       multiple sources, data model.
    3. LibreOffice recognizes it but interacts more weakly than Excel.
    """
    p = resolve_path(file_path)
    require_exists(p)
    return pivot_ooxml.create_interactive_pivot(
        p,
        source_sheet=source_sheet,
        rows=rows,
        values=values,
        agg_func=agg_func,
        columns=columns,
        page_fields=page_fields,
        filters=filters,
        target_sheet=target_sheet,
        location=location,
    )


@mcp.tool()
def list_conditional_formats(file_path: str, sheet: str | None = None) -> dict:
    """读取工作表中已有的条件格式规则（写入 + 读取闭环）。

    用途：接手一张别人的表时，先看清它埋了哪些自动规则（哪些格子会变色/变红），
    避免修改时破坏既有规则。

    返回：规则条数与明细（作用区域 / 类型 / 运算符 / 阈值公式 / 填充色 / 优先级），
    按优先级排序。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return advanced.list_conditional_formats(p, sheet)


@mcp.tool()
def compare_sheets(
    file_path: str,
    sheet1: str,
    sheet2: str,
    key_column: str,
    compare_columns: Optional[list[str]] = None,
    max_diff: int = 50,
) -> dict:
    """按关键列比对两个工作表的差异 —— 只回传差异摘要，不回传整表。

    用途：版本对比、变更检测、对账。

    返回：仅在表1 / 仅在表2 / 值有差异 三类统计，
    以及最多 max_diff 条差异明细（超出会明确说明截断数量，以统计值为准）。
    """
    p = resolve_path(file_path)
    require_exists(p)
    return advanced.compare_sheets(
        p, sheet1, sheet2, key_column, compare_columns, max_diff
    )


@mcp.tool()
def join_sheets(
    file_path: str,
    left_sheet: str,
    right_sheet: str,
    on: str,
    how: str = "left",
    columns: Optional[list[str]] = None,
    limit: int = 20,
) -> dict:
    """关联两个工作表（类似 SQL JOIN），只回传前 limit 行结果。

    参数：
    - on：关联键列名（两表都要有）
    - how：left / right / inner / outer（默认 left）
    - columns：只返回这些列（省 token），省略返回全部
    - limit：返回行数上限（默认 20）
    """
    p = resolve_path(file_path)
    require_exists(p)
    return advanced.join_sheets(
        p, left_sheet, right_sheet, on, how=how, columns=columns, limit=limit
    )


@mcp.tool()
def revert_last_write(file_path: str) -> dict:
    """回滚最近一次由本服务执行的写入（用 .bak 恢复）。

    当 write_cells 的结果不符合预期时调用它撤销。
    只保留最近一次备份，连续写入时回滚的是最后一次。
    """
    p = resolve_path(file_path)
    require_exists(p)
    bak = revert_last(p)
    return {"file": str(p), "reverted_from": bak, "status": "已回滚到上一次写入前"}


def main() -> None:
    """启动 stdio 传输（本地优先，数据不出机器）。"""
    # 与记忆系统一致：把调试输出引到 stderr，保证 stdout 是纯净的协议流
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
