# 内容 JSON Schema

`build` 子命令的 `--content` 文件格式。顶层两个键：`sectPr`（可选）与 `blocks`（必需）。

```json
{
  "sectPr": "<w:sectPr>...</w:sectPr>",
  "blocks": [ { "type": "...", ... } ]
}
```

## 顶层键

| 键 | 必需 | 说明 |
|---|---|---|
| `blocks` | ✅ | 内容块数组，按顺序渲染 |
| `sectPr` | ❌ | 页面设置。省略则用 A4 纵向默认值（与 MES 调研报告模板一致） |

默认 `sectPr`（省略时自动使用，**建议直接省略**）：

```xml
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"
         w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720" w:num="1"/>
<w:docGrid w:linePitch="360" w:charSpace="0"/></w:sectPr>
```

## blocks 各类型

### `title` — 封面主标题

```json
{ "type": "title", "text": "XX有限公司 MES 系统建设现场调研报告" }
```
居中、加粗、默认 28 半磅（=14pt）。可选 `size` 覆盖。

### `subtitle` — 封面副信息 / 落款

```json
{ "type": "subtitle", "text": "调研日期：2026年9月16日　|　编制单位：XX" }
```
居中、灰色（默认 `666666`）、默认 18 半磅（=9pt）。可选 `color`、`size` 覆盖。

### `h1` / `h2` / `h3` — 标题

```json
{ "type": "h1", "text": "一、项目基本信息" }
{ "type": "h2", "text": "2.1 公司背景" }
```
自动套用模板样式 ID：`h1`→`3`、`h2`→`4`、`h3`→`5`。
如需覆盖：`{"type":"h1","text":"...","style":"3","size":"30"}`。

### `p` — 段落

```json
{ "type": "p", "text": "正文内容，**这段会加粗**。" }
{ "type": "p", "text": "条目内容", "bullet": true }
{ "type": "p", "text": "居中文字", "align": "center" }
```

| 字段 | 默认 | 说明 |
|---|---|---|
| `text` | `""` | 支持 `**粗体**` 行内标记 |
| `bullet` | `false` | true 时行首加 `• ` |
| `bold` | `false` | 整段加粗 |
| `align` | — | `center` / `right` / `both` |
| `color` | — | 十六进制，如 `"666666"` |
| `size` | — | 半磅值，如 `"19"`（=9.5pt） |

### `table` — 表格

```json
{
  "type": "table",
  "widths": [1500, 7140],
  "rows": [
    ["项目字段", "内容"],
    ["项目名称", "XX有限公司 MES 系统建设项目"]
  ]
}
```

| 字段 | 必需 | 说明 |
|---|---|---|
| `rows` | ✅ | 二维字符串数组。首行为表头（可由 `header` 关闭） |
| `widths` | ✅ | 各列宽（twips）。**总和必须 = 8640**（A4 纵向内容区） |
| `header` | ❌ | 默认 `true`，首行加粗+底色+跨页重复 |
| `shade` | ❌ | 表头底色，默认 `"D9E2F3"` |

**列宽分配经验**（总 8640）：

| 表格类型 | 建议列宽 |
|---|---|
| 2 列 键值表 | `[1500, 7140]` |
| 3 列 产品/待办表 | `[1700, 3600, 3340]` |
| 4 列 系统/痛点表 | `[1500, 1200, 2800, 3140]` |
| 4 列 痛点汇总表 | `[1150, 3100, 700, 3690]`（痛点列最宽，优先级列最窄） |
| 5 列 仓库范围表 | `[1800, 900, 1200, 2000, 2740]` |

**单元格内换行**：字符串里写 `\n`，会拆成多个段落。

### `blank` — 空行

```json
{ "type": "blank" }
```

## 完整最小示例

```json
{
  "blocks": [
    { "type": "title",    "text": "XX有限公司 MES 调研报告" },
    { "type": "subtitle", "text": "调研日期：2026年9月16日" },
    { "type": "h1", "text": "一、项目基本信息" },
    { "type": "table",
      "widths": [1500, 7140],
      "rows": [["项目字段","内容"],["项目名称","XX MES 项目"]] },
    { "type": "h1", "text": "二、项目背景" },
    { "type": "p",  "text": "公司主营**电子组装**制造。" },
    { "type": "p",  "text": "关键要点。", "bullet": true },
    { "type": "blank" },
    { "type": "subtitle", "text": "报告编制：摩尔元数（福建）科技有限公司" }
  ]
}
```

完整可运行样例见 `examples/content_example.json`。

## 校验

生成后用 `verify` 确认结构无误：

```bash
python template_docx_report.py verify <out.docx>
```

关注输出里的「标题样式分布」——`h1` 数量应等于预期的一级章节数，`h2` 数量等二级小节数。
与 `outline` 模板骨架对照，可快速发现遗漏章节。

## 常见错误

| 现象 | 原因 | 解决 |
|---|---|---|
| 表格被拉伸/溢出 | `widths` 之和 ≠ 8640 | 调整列宽使总和为 8640 |
| 标题颜色字体不对 | 模板 styles.xml 未生效 | 确认 `--template` 指向正确模板，且模板含 `word/styles.xml` |
| `**文字**` 原样显示 | 标记符未闭合或长度不足 | 确保成对且内部非空，如 `**粗体**` |
| XML 良构校验失败 | `text` 里含未转义 `<` `>` `&` | 脚本已自动转义；若手写 XML 需自行转义 |
