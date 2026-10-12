# duduExcel 竞品分析 v1 —— GitHub Excel MCP 全景对比

> 调研日期：2026-10-12
> 调研方式：GitHub API 实测（星数/推送/许可）+ 逐仓库细读 README/工具清单
>
> 本文定位：为 duduExcel 的升级提供依据。
> **核心结论：duduExcel 的短板不是功能而是「分发」与「广度」；
> 护城河是「服务端聚合省 token + 格式语义 + 中文优先」，应放大而非丢掉。**

---

## 一、市场格局（实测星数）

| ★ | 项目 | 推送 | 许可 | 说明 |
|---|---|---|---|---|
| **4234** | **haris-musa/excel-mcp-server** | 2026-10-10 | MIT | **统治级头名** |
| 82 | jgravelle/jdatamunch-mcp | 2026-09-01 | NOASSERTION | 表格数据索引，99%+ token 节省 |
| 47 | jwadow/mcp-excel | 2026-02-28 | AGPL-3.0 | 原子操作，类 SQL |
| 20 | VOYAGER-Inc/excel-vision-mcp | 2026-07-31 | MIT | 读取 xlsx 内嵌图片 |
| 6 | senoff/xlsx-for-ai | 2026-10-11 | MIT | 50 工具，实时重算 |
| 6 | pixelsmasher13/gridpath | 2026-10-12 | NOASSERTION | MCP + 桌面应用 |
| 5 | mort-lab/excel-mcp | 2026-08-08 | MIT | openpyxl 基础 |
| 5 | mbeps/excel-mcp | 2026-10-08 | MIT | 通用 |
| 3 | zavora-ai/excel-mcp-server | 2026-08-25 | Apache-2.0 | Rust 实现 |
| 2 | eminsk/xlsx_vievers | 2026-10-11 | MIT | 129+ 函数公式引擎 |
| 1 | arttiongson/excel-live-mcp | 2026-05-18 | MIT | token 优化只读 |
| 0 | theluckystrike/mcp-spreadsheet | 2026-10-06 | MIT | 通用 |
| 0 | sahan411/spreadsheet-audit-mcp | 2026-10-04 | MIT | 公式审计 |
| 0 | madeindigio/Sheets-MCP | 2026-09-07 | MIT | Rust |
| **0** | **Shine8592/duduExcel（本项目）** | 2026-09-15 | MIT | **0★，与头名差 4234** |

**赛道判断**：极度拥挤（40+ 项目），但**只有 1 个统治级**（4234★），其余均在 100★ 以下 —— 说明这是一个"赢家通吃"的赛道，第二梯队几乎无人问津。

**这与 duduppt 当初的处境完全一致**：功能补齐了，但没人看见。

---

## 二、头名 haris-musa/excel-mcp-server 的取胜逻辑

细读其 README，强在四点：

### ① 工具广度：40+ 工具，7 大类

| 类别 | 工具 |
|---|---|
| Workbooks | `create_workbook` `describe_workbook` `set_workbook_settings` `list_workbooks` `export_workbook` `import_workbook` |
| Sheets | `describe_sheet` `create_sheet` `rename_sheet` `copy_sheet` `delete_sheet` `insert_rows_or_columns` `delete_rows_or_columns` |
| Cells | `read_range` `write_range` `clear_range` `copy_range` `sort_range` `transform_range` `find_cells` `replace_cells` |
| Formatting | `format_range` `merge_cells` `set_sheet_layout` `add_conditional_format` `add_data_validation` |
| Objects | `set_table` `create_chart` `delete_chart` **`create_pivot_table`** `delete_pivot_table` **`add_slicer`** `delete_slicer` `insert_image` `delete_image` **`add_sparklines`** `delete_sparklines` |
| Names/notes | `set_defined_name` `delete_defined_name` `set_note` `delete_note` |
| Macros | `read_vba`（写需 opt-in） |

**duduExcel 有 20 个工具** —— 广度约为头名一半。

### ② 公式引擎：内置 260 函数，无需 Excel

> 头名内置计算器（约 260 函数，用 2500+ 条真实 Excel 公式做对照测试），
> 使得 **从未在 Excel 中打开过的文件**，`read_range` 也能返回值。

**duduExcel 的 `recalculate` 依赖 LibreOffice**（未安装则降级）。这是真实差距。

不过要客观：duduExcel 的 `recalculate` 是**真实重算**（LibreOffice 引擎），而头名是"内置近似计算"，且头名也承认"无法精确复现时返回 null 并列在 uncalculated 中，绝不猜"。两者路线不同，各有取舍。

### ③ 高级图表与对象

- 基础 8 种：柱/条/线/面积/饼/环/雷达/散点/气泡 + 组合、次坐标轴、趋势线、误差线
- **Excel 2016 类型**：瀑布图、直方图、帕累托、箱线、树图、旭日、漏斗
- **切片器 + 时间轴**（透视表交互）
- **迷你图 sparklines**

**duduExcel**：`add_chart` 基础图表；透视表是亮点（手写 OOXML 可交互）。但缺 2016 图表、切片器、迷你图。

### ④ 工程与分发

| 项 | 头名 | duduExcel |
|---|---|---|
| 一键运行 | **`uvx excel-mcp-server stdio`**（uv，零安装） | `pip install duduexcel` |
| 安装文档 | 官网 excelmcpserver.com + 各客户端示例 | README 有 |
| MCP bundle | `.mcpb` 一键装（Claude Desktop） | ❌ |
| HTTP 远程 + Docker | ✅ | ❌ |
| 大文件流式 | ✅ 百万行内存平坦 | ❌ |
| 官网 | ✅ | ❌ |

**分发是最大差距**：`uvx` 让头名"零安装就能跑"，这是病毒式传播的关键。

---

## 三、duduExcel 的真实护城河（必须放大）

对照头名，duduExcel 有三项**头名没有**的能力：

### ① 服务端聚合 —— 省 token（核心差异）

```python
filter_count()   # 只返回"有多少条满足"，不返回数据行
aggregate()      # 分组聚合，每组一个数字
top_n()          # 排行榜，只回传 TopN
join_sheets()    # 表关联
compare_sheets() # 表差异
```

头名的 `read_range` 是把数据取回来；duduExcel 是**在服务端算完只回结果**。对大表，这是数量级的 token 节省。

> 注：`jdatamunch-mcp`（82★）也主打 token 节省（99%+），说明这个方向有真实需求，但它是"索引 CSV/Excel 后查询"，与 duduExcel 的"直接对 xlsx 聚合"路线不同。

### ② 格式语义 —— 读懂作者留下的痕迹

```
[B] 粗体 / [I] 斜体 / [S] 删除线（常表示"已取消/作废"）
[HL:色] 底色高亮（常表示"待审阅/重点"）
[M] 合并单元格 / [HIDDEN-REF] 隐藏但被公式引用
```

**头名完全没有这个概念** —— 它把格式当作"样式"，duduExcel 把格式当作"语义"。
这是独到的洞察（汲取自 excel-vision-mcp 的思路）。

### ③ 中文优先

- `apply_chinese_style`：微软雅黑 + 自动列宽（中文按 2 字符宽）
- 人民币格式 `¥#,##0`、零显示短横 `¥#,##0;(¥#,##0);-`
- 中文表格自动美化

**头名无任何中文特化。**

---

## 四、差距清单（按优先级）

### 🔴 P0 —— 分发（和 duduppt 一样，这是 0★ 的真因）

| 项 | 说明 |
|---|---|
| **uvx 支持** | 让 `uvx duduexcel` 可直接运行（对标头名的零安装体验） |
| **PyPI 元数据完善** | 当前已发布 0.3.0，但需完善 keywords/classifiers、补 `mcp-name` 标记 |
| **MCP 目录收录** | glama.ai / mcp.so / pulsemcp / awesome-mcp-servers |
| **git tag 为空** | 从未打过 tag / 从未发过 Release —— 版本不可追溯 |

### 🟡 P1 —— 广度补齐（不追求全量，挑高价值）

| 能力 | 头名有 | duduExcel | 价值 |
|---|---|---|---|
| 数据验证（下拉）| ✅ | ❌ | 高（表单类表格刚需）|
| 排序/查找替换/去重 | ✅ | ❌ | 高（清洗刚需）|
| 迷你图 sparklines | ✅ | ❌ | 中 |
| 切片器 | ✅ | ❌ | 中（依赖透视表）|
| Excel 2016 图表 | ✅ | ❌ | 中 |
| 定义名称 / 批注 | ✅ | ❌ | 低 |
| 大文件流式读 | ✅ | ❌ | 高（大表场景）|

### 🟢 P2 —— 放大护城河

- **聚合能力增强**：更多聚合算子、多列分组、透视式交叉表
- **格式语义增强**：把 `[S]` 删除线等做成可查询/可统计（如"统计所有作废行"）
- **中文增强**：中文日期/数字格式、中文排序、中文列宽智能适配

---

## 五、升级路线

```
第一步：分发先行（最高性价比，和 duduppt 同一策略）
  ├─ uvx 支持（pyproject 加 [project.scripts] + 确保 uv 可跑）
  ├─ 完善 PyPI 元数据
  ├─ 打 tag + 发 Release（当前 tag 为空！）
  └─ 提交 MCP 目录收录

第二步：放大护城河（差异化立住）
  ├─ 聚合算子增强
  └─ 格式语义可查询化

第三步：广度补齐（挑高价值的做）
  ├─ 数据验证
  ├─ 排序/查找替换/去重
  └─ 大文件流式读

第四步：版本对齐 + 最终核查
```

---

## 六、一句话战略

**不要全面对标头名的 40+ 工具**（它有先发优势和 4234★ 生态）。

duduExcel 应该打：
> **「省 token 的中文 Excel 助手」** —— 服务端聚合让你不烧钱，
> 格式语义让你读懂作者的意图，中文优先让你交付即合规。

它卷"功能全"，你卷"**省钱 + 懂中文 + 懂格式**"。
