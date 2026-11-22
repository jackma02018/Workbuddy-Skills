---
name: 摩尔多页PPT合并
description: "把多份单页海报式方案 PPT 合并进一份模版 PPT，并完成三步加工：①按模版统一为黑体、标题20号/正文12号；②把每页内容整体合并为一个对象（组）后等比缩小；③文字自适应容器（缩字号或减字数，不溢出、不压 Logo）。合并与缩放过程中强制保留每页标题文字与页面右侧的摩尔元数公司图标，且始终置顶可见。内置 4 个脚本（合并 / 分组缩放 / 文字自适应 / 完整性校验）与矢量图无损复制方案，可绕过 PIL 对 EMF/WMF/SVG 的识别失败。触发场景：把几份单页方案拼成一份模版、多页 PPT 合并、合并后整体缩小、整页内容合并成一个对象、标题或正文超出容器或压住右上角 Logo、合并后图片丢失或字体没统一、PPT 方案合集交付。"
version: 1.0.0
display_name: 摩尔多页PPT合并
display_name_en: Moore Multi-page PPT Merge
disable-model-invocation: false
agent_created: true
---

# 摩尔多页PPT合并

把「N 份单页海报式方案 PPT + 1 份公司模版」加工成「一份字体统一、每页可整体缩放、文字不溢出、标题与 Logo 始终可见」的方案合集。

核心价值：**图片一张不丢、标题与公司图标永远置顶可见、文字永远在容器内。**

## 何时使用

- 「把这几份方案合并到《摩尔元数方案模版.pptx》」「多页 PPT 合并成一份」
- 「每页内容整体合并成一个对象，再整体等比缩小」
- 「缩小/合并时必须保留标题文字和右上角的摩尔元数 Logo」
- 「文字超出容器了 / 压住 Logo 了，自动缩小字号或减少字数」
- 「合并后图片丢了 / 字体没统一 / 打开报错」

## 铁律（必须遵守）

1. **先备份再改**。任何修改前复制一份 `原名_<阶段>前备份.pptx`，贯穿全流程可回溯。
2. **图片一律走原始字节复制**，禁止用 `get_or_add_image_part`（内部依赖 PIL，遇到 EMF/WMF/SVG 直接抛 `cannot identify image file`，图片会破）。用 `Part(partname, content_type, package, blob)` + `relate_to(target, RT.IMAGE)`，详见 `references/pptx-internals.md`。
3. **标题与右侧公司图标不参与合并、不参与缩放**，且必须移到 spTree 末尾（PPT z 序 = XML 文档顺序，末尾即最上层），保证不被组遮挡。
4. **改完必验**：图片 sha1 逐张比对 + 组结构 + 标题/Logo 是否在组外顶层 + 文字是否仍在容器内。禁止「改完就交付」。
5. **本机无 PowerPoint/WPS**：无法做 Office 端像素级验收。涉及观感结论时必须如实说明「未做 Office 端渲染验收」。
6. **环境注意**：本机 bash 被 shim 破坏（`dirname`/`cd`/`ls` 不可用），一律用绝对路径调用 Python：`C:\Users\majunpeng\.workbuddy\binaries\python\versions\3.13.12\python.exe`。

## 工作流程（4 步，可单独执行）

```
Step 0  备份                     →  复制 模版.pptx → 模版_合并前备份.pptx
Step 1  合并 + 统一字体            →  scripts/merge_ppts.py
Step 2  整页合并为一个对象 + 等比缩小  →  scripts/group_and_scale.py
Step 3  文字自适应容器              →  scripts/fit_text.py
Step 4  完整性校验                 →  scripts/verify_pptx.py
```

### Step 1 — 合并并统一字体

```bash
python merge_ppts.py --template 摩尔元数方案模版_原单页备份.pptx \
  --sources A.pptx B.pptx C.pptx D.pptx \
  --out 摩尔元数方案模版.pptx --font 黑体 --title-pt 20 --body-pt 12
```

- 模版保留首页作封面（`--keep first`），每份来源取其第 1 页追加。
- 字号规则：原始字号 ≥ `--title-threshold`（默认 16pt）→ 统一 `--title-pt`；其余 → `--body-pt`。
- 中文字体必须同时写 `a:latin` / `a:ea` / `a:cs` 的 `typeface`，否则只改了西文字体。

### Step 2 — 整页合并为一个对象 + 等比缩小

```bash
python group_and_scale.py --file 摩尔元数方案模版.pptx --scale 0.85 --skip-slides 1
```

- 每页把「非标题、非 Logo」的全部形状移进**一个** `p:grpSp`；标题与右侧 Logo 留在顶层。
- 缩放实现：组的 `a:ext`（屏幕尺寸）= `scale × a:chExt`（原始坐标系），`a:xfrm` 位于 `p:grpSpPr` **内部**，取用 `.//a:xfrm`。
- `--scale` 建议 0.85；脚本按 `min(可用宽/页宽, 可用高/页高, 0.9)` 自动收敛，`--margin-left/--margin-top` 默认 0.3in / 0.8in。
- 组缩放会连带缩放组内文字，因此组内文字与容器始终同比例，不会失配。

### Step 3 — 文字自适应容器

```bash
python fit_text.py --file 摩尔元数方案模版.pptx --min-pt 10 --logo-gap 0.2
```

对**顶层标题区文字**（标题、副标题）逐条处理，顺序固定：

1. 容器宽度封顶：右缘不超过「最左 Logo 左缘 − `--logo-gap`」，避免压住公司图标。
2. 高度保底：至少容纳 `1.25 × 最大字号 + 上下内边距`，消除「假溢出」。
3. 横向溢出（`wrap="none"`）：按比例缩小字号，下限 `--min-pt`（默认 10pt）。
4. 仍溢出 → **减少字数**：从尾部逐字删减并补 `…`，直到入框。
5. 写入 `a:normAutofit`：以后在 PowerPoint/WPS 里改字，溢出时自动缩小。

> 组内正文不要批量改字号：原始海报在设计尺寸下本就贴合，整组已同比例缩放；盲目批量缩字会破坏设计。确有溢出时按 `references/pptx-internals.md` 的「精准定位法」单独处理。

### Step 4 — 校验（不可跳过）

```bash
python verify_pptx.py --file 摩尔元数方案模版.pptx --baseline 摩尔元数方案模版_合并前备份.pptx
```

输出 5 项结论：图片 sha1 是否零缺失零变更（排除 zip 目录占位项）、每页组数与缩放比、标题/Logo 是否在组外顶层、文字估算是否仍在容器内、`normAutofit` 是否已写入。

## 输入参数

| 参数 | 说明 | 默认 |
|---|---|---|
| `--template` | 公司模版 pptx（首页作封面） | 必填 |
| `--sources` | 一份或多份单页方案 pptx | 必填 |
| `--out` | 合并输出路径 | 必填 |
| `--font` / `--title-pt` / `--body-pt` | 字体与字号 | 黑体 / 20 / 12 |
| `--title-threshold` | 视作标题的原始字号门槛 | 16 |
| `--scale` | 整页等比缩放比 | 0.85 |
| `--skip-slides` | 前 N 页不参与分组缩放（封面） | 1 |
| `--min-pt` | 文字缩放下限 | 10 |
| `--logo-gap` | 标题右缘与 Logo 左缘的最小间距 | 0.2in |

## 扩展场景：空白页按封面版式重建

用户插入空白页并要求「参考第 1 页布局」时：

1. 把空白页的版式关系重定向到封面所用 `slideLayout`（先 `del part.rels._rels[rId]`，再 `part.relate_to(layout_part, RT.SLIDE_LAYOUT)`）——绿点、弧线、Logo 等装饰来自版式，会自动出现。
2. 深拷贝封面页的标题形状到该页，再交给 `fit_text.py` 做自适应。

详见 `references/pptx-internals.md`「空白页复刻封面版式」。

## References

- `references/pptx-internals.md` — OOXML/python-pptx 深水区笔记：矢量图复制、rId 重映射、组变换语义、`Pt()` 陷阱、zip 目录占位项、z 序、版式重定向。
- `references/prompt-templates.md` — 下次可直接复用的提示词模板（含参数填写位）。
- `scripts/` — `merge_ppts.py` / `group_and_scale.py` / `fit_text.py` / `verify_pptx.py`。

## 交付时主动说明

1. 图片完整性比对结果（张数、是否零缺失）、每页缩放比；
2. 标题与右侧 Logo 是否确认在组外顶层；
3. 有文字被缩小或截断时，逐条列出「原字号 → 新字号」或「截断后文本」；
4. 本机无 Office，未做 PowerPoint/WPS 端渲染验收。
