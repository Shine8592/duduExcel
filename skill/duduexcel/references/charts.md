# 图表做法与保真度

> 这是 duduExcel 的差异化能力：Anthropic 官方 xlsx skill 正文无图表指导，
> 竞品 knorq 的 Known Limitations 明确写着不支持图表。

## 支持类型

| chart_type | openpyxl 类 | 适用 |
|---|---|---|
| `bar` | BarChart | 类别对比（最常用） |
| `line` | LineChart | 时间趋势 |
| `pie` | PieChart | 占比（类别 ≤ 7 个） |
| `scatter` | ScatterChart | 相关性 |

## 调用要点

```
add_chart(
  file_path="...",
  sheet="销售",
  data_range="C2:C7",          # 数值区域
  categories_range="A2:A7",     # 分类轴（X 轴标签）
  chart_type="bar",
  title="月度销售额",
  anchor_cell="E2"              # 图表左上角锚点
)
```

- **区域字符串用纯区域即可**（如 `C2:C7`），服务内部会正确解析，
  不需要写成 `表名!C2:C7`（中文表名还需引号，易错）。
- `anchor_cell` 建议放在数据区右侧或下方，避免盖住数据。

## 保真度注意事项（重要）

1. **openpyxl 写入的图表不含渲染结果**，只有定义。
   Excel/LibreOffice 打开时会按定义渲染，通常正常。
2. **执行 `recalculate` 后请重开文件确认图表仍在**。
   LibreOffice 转换过程对部分图表属性（如自定义配色）可能做归一化。
3. 饼图建议类别 ≤ 7 个，否则切片过碎难以阅读。
4. 中文标题一般没问题，但若出现乱码，改用**微软雅黑**。

## 图表 + 中文的搭配

- 先 `apply_chinese_style` 美化数据区
- 再 `add_chart` 插入图表
- 分类轴若为中文，`apply_chinese_style` 的自适应列宽会让标签更易读

## 尚不支持

- 透视表（pivot table）
- 条件格式（conditional formatting）
- 图表深度定制（次坐标轴、组合图、自定义配色序列）

这些能力计划在后续版本补齐；需要时可用 Python + openpyxl 直接写脚本。
