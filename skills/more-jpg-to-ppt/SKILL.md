---
name: more-jpg-to-ppt
description: >-
  把图片转成 PPT，支持两种输出模式——图片版（一页一图，忠实保真）和矢量版（1:1 重绘为可编辑形状与文字）。
  当用户发来一张或多张图片（截图、架构图、流程图、示意图、方案图），并要求"做成 PPT / 转成 PPT / 画成 PPT /
  转成可编辑的 PPT / 我要能改文字的 PPT"时使用。单张图默认输出 1 页。
agent_created: true
---

# 图片转 PPT（双模式）

## 结论先行：先判模式，再动手

拿到图片后**第一个动作不是转换，是判模式**。判错的代价是返工重做。

| 判据 | 走哪条路 |
| :--- | :--- |
| 用户明确说"要能改文字 / 可编辑 / 矢量" | → **矢量版** |
| 用户明确说"原样放进去就行 / 忠实还原 / 不用改" | → **图片版** |
| 图是**架构图 / 流程图 / 系统方案图**，且有改字需求 | → **矢量版** |
| 图是**照片 / 复杂截图 / 数据大屏 / 界面截图** | → **图片版**（矢量重绘照片不现实） |
| **多张图**且要求"一页一图" | → **图片版** |
| 图片**不是 16:9**，用户要"满页铺满" | → **图片版**（矢量版按 1280×720 重排，比例不符会留白或变形） |
| 说不清 | → **两个版本都出**，文件名带 `_图片版` / （矢量版不带后缀） |

通常单张架构图 → 矢量版更受用；一批截图 → 图片版。

## Mode A：图片版（一页一图）

### 做法

1. 把待转图片**复制到一个独立临时文件夹**（不要直接指向混杂目录，脚本只处理顶层、不递归，混入其它图会一起转进去）。
2. 跑 `scripts/img2ppt.py`：

```bash
python scripts/img2ppt.py <图片文件夹> <输出.pptx> [--size 16:9|4:3] [--background FFFFFF] [--margin 0]
```

3. **用完删掉临时文件夹**，不要留在用户工作目录里。
4. 校验：页数 = 图片数；每页恰好 1 张图；顺序 = 文件名自然顺序（`img2` 在 `img10` 之前）。

### 行为约定

- 图片始终**等比放大铺满**页面（等价 PowerPoint「适应幻灯片」），不做 1:1 限制、不裁切。
- 支持 png / jpg / jpeg / bmp / gif / webp / tif / tiff；webp 自动转 PNG 后插入。
- 以 `.` 开头的隐藏文件跳过；损坏图片打警告跳过该页，不中断整体转换，完成后汇报跳过了哪些。
- 依赖 `python-pptx` + `pillow`。

## Mode B：矢量版（可编辑单页）

把原图当**设计稿**，逐区域 1:1 重绘到 slidep DSL，产出 1280×720 单页 PPTX，所有文字/色块/图标可在 PowerPoint/WPS 里直接改。

### 步骤

1. **读图并记录**（这一步决定后面像不像，别省）：
   - 整体布局：垂直分层 / 水平分栏 / 中心辐射 / Z 型阅读路径
   - 分区与边框：虚线框/实线框/圆角框，各自的标签
   - 配色：逐个色块取近似 hex
   - 图标：每个图标对应什么语义
   - 文字：所有标题、标签、列表项、脚注的**逐字文本**
   - 连接关系：箭头方向、线型、颜色
2. **加载 `tencent-pptx` skill**，按 `references/create-from-scratch.md` 执行。
3. **建项目结构**：
   ```
   <workspace>/<项目目录>/
   ├── STORY.md          # 视觉叙事逻辑（各区域内容）
   ├── DESIGN.md         # 色板 hex / 字体 / 布局尺寸 / 元素清单
   ├── slides/01.slide   # 单页源文件
   └── <中文名>.pptx
   ```
   目录名 = pptx 文件名去后缀。
4. `slidep create "<项目目录>/<中文名>.pptx"` 建空白文件，**立即 `present_files` 打开预览**（后续逐页制作时能实时看效果）。
5. **写 `slides/01.slide`**：1280×720 画布，`position: 'absolute'` + 精确 top/left/width/height。
6. `slidep lint` 校验 → 修到 `ok:true`（**诊断要当视觉检查读，不只为过检**）→ `slidep upsert-dsl ... --page-index 0` 提交。
7. **交付前核页数**：`upsert-dsl` 是追加语义，重复提交会堆成多页。用 python-pptx 确认 `len(prs.slides)` 是 1。
8. `present_files` 交付。

> ⚠️ **单页图必须带 `--page-index 0`**。`slidep create` 出来的文件自带一页空白页，不带参数会追加成两页。若已堆多页，slidep 没有删页命令，只能删掉 pptx 重建。

### 动笔必守的硬约束

**DSL 是受限 DSL，不是 React。** 把它当 HTML + 内联 CSS 写。

- 禁止 `import` / `export` / `module` / hooks / state / 函数组件定义；文件最后一个表达式必须是 `<Slide>`。
- `Box` 只支持 `display: 'flex'`——**禁止 grid**，**不支持 `calc()`**。
- `box-sizing: border-box`：`width` 已含 padding 和 border。
- 高度必须自己算：`上padding + 标题区 + 内容区 + 间距 + 下padding ≤ 720`。**绝对不允许溢出。**
- Slide 的 `padding` 是安全区，不能被绝对定位绕过。
- 换行用 `<br />`，不用 `\n`；局部样式用 `<span style={{...}}>`，不支持 `<strong>`/`<em>`/`<u>`。
- 图标优先 `FAIcon`（**`fill` 不设则完全不可见**，且图标名必须真实存在，禁编造），特殊图形/插画用内联 `<svg viewBox>`（**`<svg>` 内嵌 `<text>` 可渲染**）。
- 文案里出现 `{` 或 `}` 必须包成 `{'...'}`，否则 JSX 当表达式解析报错。
- 长文本加 `whiteSpace: 'nowrap'` 防非预期折行。
- 禁止自创配色——颜色取自原图。禁止省略原图信息——所有模块名、列表项、图例文字逐字复刻。

### 绝对定位箭头组件的正确形态

这是最容易踩的坑。**标签文字和箭头 SVG 拆成两个绝对定位 `Box`**，不要塞进同一个 flex column（文字会被挤到溢出）：

```
<Box style={{ position:'absolute', top:96, left:466, width:14, height:76, ... }}>
    <svg width={14} height={76} viewBox='0 0 14 76'> ...虚线+三角箭头... </svg>
</Box>
<Box style={{ position:'absolute', top:96, left:484, width:200, height:76,
              flexDirection:'column', gap:4, justifyContent:'center' }}>
    <Text style={{ fontSize:15, fontWeight:'bold' }}>数据库同步</Text>
    <Text style={{ fontSize:12 }}>（如 Oracle 主备同步）</Text>
</Box>
```

带上下文字标注的横向箭头同理：`Box` 高度取 `字号×1.4 + 箭头高 + 2`，一般 36~40px。

### 布局自检清单（提交前逐条过）

- [ ] 总高度算过，内容不溢出 720
- [ ] **分区卡内高度算过**：卡头(44) + 卡体内容 ≤ 卡总高，纵向堆叠的内嵌卡每张按 `padTop + 标题行 + 行数×行高 + 间距 + 胶囊` 手算
- [ ] **骑缝胶囊标签**（压在卡顶边的「主库（写）」这类）要**移出卡片**做独立绝对定位 Box，否则被 `overflow:'hidden'` 裁掉
- [ ] **窄卡长文案去掉 `whiteSpace:'nowrap'`**，宽度 ≈ 字数 × 字号，算不下就让它折行
- [ ] **等宽/固定宽小格里的长内容**（表头格、标题栏、页码位）先算剩余宽再放：宽 124 的 4 列表，前 3 列各 34 只剩 22 给第 4 格，`2024-001` 必然溢出。**解法是重新分配列宽**（标签列压到 26，数值列给足 62），不是缩字号硬塞
- [ ] **图标/徽标/示意图等造型元素，动手前裁原图放大 3 倍看过**，确认实心还是线框、几层色阶、外轮廓形状，再写 SVG（lint 查不出造型偏差，只有回看才暴露）
- [ ] **同一张图里同一"角色"的元素，逐个裁开单独确认**，不要"看一次推广到全部"。原图常有细节差异：同一个机器人，圆底头像版本是"双层机身+耳翼+天线+大椭圆眼+嘴缝"，大方图标版本是"单层机身+折角"，品牌位版本是"浅蓝在上深蓝在下"。按一个改，另一个必错
- [ ] **改通用绘制函数的签名/路径前，先 `grep -n "函数名("` 列出全部调用点**。通用 `robot()` / `avatar()` 被多处复用，改机身几何会连带改坏用同一函数画的其它图标（本次改机器人把①的用户头像改成了"棺材板"）。要么把角色拆成两个独立函数（`robotAvatar()` / `userAvatar()`），要么给每个调用点单独传形状参数
- [ ] 每个视觉区块有独立容器 Box，必要时加背景增强分区感
- [ ] 原图所有文字逐字复刻，无省略
- [ ] 配色全部取自原图，未自创
- [ ] 图标语义可辨识
- [ ] lint 通过且诊断已当视觉检查读过一遍（见下）
- [ ] 自渲染预览图与源图逐分区比对过（见 `references/layout-verify.md`）

> 📌 **横排 4 卡 / 通栏卡的布局**见 `references/vector-mode.md` 的「4 分区横排布局」小节，含高度预算模板与内嵌卡溢出对策。

## 版本并存惯例

矢量版和图片版**都保留**，文件名区分，让用户按场合挑：

- `XXX.pptx` —— 矢量版（要改内容时用）
- `XXX_图片版.pptx` —— 图片版（要绝对保真时用）

## 备份自己的技能包（可选收尾）

技能装在 `~/.workbuddy/skills/` 下，**重装系统会被清空**。若用户担心丢失，可打包成 zip 备份到资料库。

打包内容：`SKILL.md` + `references/` + `scripts/` + 一份 `恢复说明.md`（写明放回路径与依赖）。

**注意资料库对 `.zip` 的路由**（踩过的坑）：

| 容器 | 结果 | 原因 |
| :--- | :--- | :--- |
| `drive` 网盘 | ❌ 被拒 | 后端把 `.zip` 路由给 page 模块 |
| `page` 页面 | ❌ 不适用 | 要求包内至少有一个 `.html` 入口 |
| `attachment` 节点附件 | ✅ 可用 | 无类型限制，挂到某文档节点下 |

正解是走**附件**：

```bash
python "<library>/attachment/attachment.py" upload --token-stdin \
  --node-id <所属文档nodeId> <技能包.zip> --file-name "技能包_YYYYMMDD.zip"
```

成功输出 `KS_ATTACHMENT_UPLOAD_OK`，再用 `download` 回读链接验收。

> 资料库 doc 的正文编辑有组件语法限制（`content` 禁裸 Markdown；`<Mark>` 无 `code` 属性；`<Table>` 不收任何属性；`--actions-file` 要传 Windows 路径）。细节见 `references/library-backup.md`。

## 已知环境坑

- **本机 bash 缺 `/usr/bin` 在 PATH 里**，直接跑 `ls` / `cp` 会 command not found。先执行：
  ```bash
  export PATH="/usr/bin:/bin:/c/Windows/System32:/c/Windows:$PATH"
  ```
- slidep 在 Node 托管环境：`<node>/node_modules/@tencent/slidep`，可执行文件 `<node>/slidep`。
- Windows 路径给 slidep 时用**正斜杠 `/`**，避免解析错误。
- `slidep screenshot` 在本机不可用（见 `references/layout-verify.md`）。
- 不要用 `reg.exe` 探测 Office 安装——常被沙箱程序黑名单拦截。直接 `find` 找 `POWERPNT.EXE` / `wpp.exe`。
