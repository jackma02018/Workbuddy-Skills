# 矢量版：图片 → 可编辑单页 PPT

把原图当设计稿，用 slidep DSL 1:1 重绘。产出 1280×720 单页，所有文字/色块/图标可在 PowerPoint/WPS 里直接改。

## 前置

先加载 `tencent-pptx` skill，按其 `references/create-from-scratch.md` 执行。本文件只补充"图片转矢量"特有的部分。

## 读图记录清单

**这一步决定成品像不像，别省。** 逐项写下来再动笔：

1. **整体布局** —— 垂直分层 / 水平分栏 / 中心辐射 / 网格排列 / Z 型阅读路径
2. **分区与边框** —— 虚线框/实线框/圆角框，各自的标题标签（如"整体架构""配置策略"）
3. **配色** —— 逐个色块取近似 hex，记主色/辅色/强调色
4. **图标** —— 每个图标对应什么语义（数据库 / 扫码枪 / 齿轮 / 时钟 / 灯泡 / 警告三角…）
5. **文字** —— 所有标题、标签、列表项、脚注的**逐字文本**（含标点）
6. **连接关系** —— 箭头方向、线型（实线/虚线）、颜色、箭头上的标注文字
7. **图例** —— 底部或侧边的说明区

## 项目结构

```
<workspace>/<项目目录>/
├── STORY.md          # 视觉叙事逻辑：从上到下/从左到右各区域讲什么
├── DESIGN.md         # 色板 hex / 字体 / 布局尺寸(基于1280×720) / 元素清单
├── slides/01.slide   # 单页源文件
└── <中文名>.pptx
```

目录名 = pptx 文件名去后缀。

## 流程

```bash
slidep create "<项目目录>/<中文名>.pptx"
# 立即 present_files 打开预览，后续制作时能实时看效果

# 写 slides/01.slide ...

slidep lint "<项目目录>/<中文名>.pptx" --dsl-file "slides/01.slide"
slidep upsert-dsl "<项目目录>/<中文名>.pptx" --dsl-file "slides/01.slide"
```

- **写完立即 lint**（硬约束）。未过则改 `.slide` 重跑，直到 `ok:true`。
- **过了立即 upsert**（硬约束）。返回 `{ok:true}` 才算提交成功。
- Windows 路径给 slidep 时用正斜杠 `/`。

### ⚠️ upsert-dsl 是「追加」语义，不是「更新」（踩过）

**第一次 `create` 出来的 pptx 本身已自带一页空白页。** 直接 `upsert-dsl` 会**再追加一页**，重复提交就累积成多页。

| 场景 | 命令 |
| :--- | :--- |
| 单页图，第一次写入 | `upsert-dsl ... --page-index 0`（覆盖那页空白页） |
| 改已有单页 | `upsert-dsl ... --page-index 0` |
| 多页追加新页 | 省略 `--page-index` |

**slidep 没有删页命令**。如果已经误堆成多页，唯一办法是 `rm` 掉 pptx、`create` 重建、再 `upsert-dsl --page-index 0`。

> 交付前务必用 python-pptx 核一遍 `len(prs.slides)`，确认页数符合预期。

## 高度预算

画布固定 1280×720，**内容绝对不允许溢出**。动笔前先算：

```
上padding + 标题区高度 + 内容区高度 + 各间距(gap/margin) + 下padding ≤ 720
```

推荐分区（本类方案图验证有效）：

| 区域 | y 范围 | 说明 |
| :--- | :--- | :--- |
| 顶部并列卡 | 8–88 | 问题 / 对策 左右对位 |
| 中部主区 | 96–420 | 左 2/3 主体 + 右 1/3 侧栏 |
| 底部区 | 428–700 | 左 2/3 横排步骤 + 右 1/3 补充说明 |

全宽 1280，左右外边距 16 → 内容宽 1248。

### 变体：4 分区横排布局（1 行 4 卡 / 2 行通栏）

源图是「4 张卡横向并列 + 1 张通栏卡」时（常见于 5 模块方案图）：

```
页头：    y   0 – 70    通栏深蓝，主标题 + 副标题居中
第一行：  y  80 – 418   4 张卡并列，卡间距 14，卡宽 = (1248 - 3×14) / 4 ≈ 300
第二行：  y 428 – 628   1 张通栏卡 w 1248
页脚条：  y 638 – 686   通栏浅蓝底
```

**卡内高度预算模板**（卡总高 338，卡头 44 → 卡体 296）：

| 用量 | 高度 |
| :--- | :--- |
| 正文区（3~4 行 13.5px，行高 1.62） | 88 ~ 110 |
| 内嵌子卡 | 148 ~ 168 |
| **合计上限** | **≤ 296，余量留 8~30** |

> ⚠️ **纵向堆叠的内嵌卡是高度杀手。** 三张 84px 的卡 + 2 个 6px gap = 265px，
> 一旦起点超过 30 就必然溢出，缩字号治不了本。
> **对策：把"状态类"信息从卡里抽出来做成通栏状态胶囊条**（高度 22），
> 只保留 2 张正文卡，纵向需求立刻从 265 降到 146。
> 判断依据：凡是**一句话就能说完**的结论（如「一次迁移，长期受益」「配置即可生效」），
> 都不该占一张卡 —— 合并成胶囊条反而更醒目。

### 跨列文字的宽度死结

横向并列卡的宽度固定（≈300，减 padding 后可用 ≈278）时，
**11.5px 中文约 11.5px/字**，一行最多约 24 字。超过就必须折行：

| 文案长度 | 处理 |
| :--- | :--- |
| ≤ 12 字 | 单行 `whiteSpace:'nowrap'` |
| 13 ~ 24 字 | 单行，但要去掉 `nowrap` 让它能折 |
| > 24 字 | **必须允许折行**（去掉 `nowrap` + 显式 `textAlign:'center'`） |

踩过的坑：备库卡只给 100px 宽，里面塞「报表分析 / 数据导出」（11.5px，约 104px）
—— lint 报 `left+2px, right+2px` 横向溢出。**窄卡里的长文案去掉 `nowrap`，让它自己折。**

### `lineHeight` 会触发 lint 假溢出（必踩，优先排查）

`<Text>` 上写了 `lineHeight`，lint 就会报 `no-overflow` + `child-containment`
的 `[bottom]` 巨大溢出（几百至上千 px，如 `Bottom+888px`）。
**这是测高 bug，不是真的挤了**，跟 width/height/flex 都无关。

同一 Box 只改这一个变量的对照实测：

| 变体 | 结果 |
| :--- | :--- |
| 单行 + 有 `lineHeight` | FAIL |
| 两行 + 有 `lineHeight` | FAIL |
| 不设 height + 有 `lineHeight` | FAIL |
| 单行 / 两行 + **无** `lineHeight` | PASS |
| 加了 `display:'flex'` + 有 `lineHeight` | 仍 FAIL（flex 救不了） |

**修法：全文不写 `lineHeight`。** 多行文案改成「一行一个 Box + 绝对定位 top」，
用 top 差值精确表达行距，视觉与原图完全等价：

```
<Box style={{ position:'absolute', left:30, top:393, width:291.5, height:26 }}>
    <Text style={{ fontSize:15.5, color:'#3D4E63', textAlign:'center' }}>供应商发货不及时，</Text>
</Box>
<Box style={{ position:'absolute', left:30, top:419, width:291.5, height:26 }}>
    <Text style={{ fontSize:15.5, color:'#3D4E63', textAlign:'center' }}>物料无法按时到位。</Text>
</Box>
```

改完全文搜 `lineHeight`，命中数必须为 **0** 再 lint。
注意 `<br />` 有两种写法，批量拆分时两种都要覆盖：
`<br />` 结尾换行、`<br />` 内联在同一行（如 `甲<br />乙`）。

### Windows 下 slidep 的真实调用方式

`slidep` 是 `/bin/sh` 包装脚本，**不能用 `node slidep ...` 调**（会报
`SyntaxError: missing ) after argument list`）。两种可行写法：

- 直接执行：`"$HOME/.workbuddy/binaries/node/versions/22.22.2-3/slidep" lint ...`
- 绕开壳（更稳，推荐）：
  `node "$BASE/node_modules/@tencent/slidep/dist/index.js" lint ...`
  其中 `BASE=~/.workbuddy/binaries/node/versions/22.22.2-3`

## 分区写法模板

**带标题条的分区容器**（原图那种蓝色标题条 + 白底描边卡）：

```
<Box style={{ position:'absolute', top:96, left:16, width:820, height:324,
              background:'#FFFFFF', border:'1px solid #C9DDF2',
              borderRadius:8, overflow:'hidden' }}>
    <Box style={{ height:30, background:'#1A7FE0', justifyContent:'center', paddingLeft:12 }}>
        <Text style={{ fontSize:16, fontWeight:'bold', color:'#FFFFFF' }}>整体架构</Text>
    </Box>
    <Box style={{ position:'relative', flex:1 }}>
        {/* 分区内用绝对定位精确摆放 */}
    </Box>
</Box>
```

注意 `overflow:'hidden'` 让标题条的四角跟随圆角。分区内层用 `position:'relative'` + 子元素 `position:'absolute'` 精确定位。

## 箭头组件正确形态（易踩坑）

**标签文字和箭头 SVG 拆成两个独立绝对定位 Box**，不要塞进同一个 flex column —— 文字会被挤到溢出。

```
{/* 垂直线 + 双向三角 */}
<Box style={{ position:'absolute', top:96, left:466, width:14, height:76,
              flexDirection:'column', alignItems:'center', justifyContent:'center', gap:4 }}>
    <svg width={14} height={76} viewBox='0 0 14 76'>
        <line x1='7' y1='4' x2='7' y2='34' stroke='#2E7FD4' strokeWidth='2' strokeDasharray='5 4' />
        <polygon points='7,0 2,10 12,10' fill='#2E7FD4' />
        <line x1='7' y1='42' x2='7' y2='72' stroke='#2E7FD4' strokeWidth='2' strokeDasharray='5 4' />
        <polygon points='7,76 2,66 12,66' fill='#2E7FD4' />
    </svg>
</Box>

{/* 标签另起一个 Box，放在箭头右侧，留足宽度 */}
<Box style={{ position:'absolute', top:96, left:484, width:200, height:76,
              flexDirection:'column', gap:4, justifyContent:'center' }}>
    <Text style={{ fontSize:15, fontWeight:'bold', color:'#14456F' }}>数据库同步</Text>
    <Text style={{ fontSize:12, color:'#7B8A9A' }}>（如 Oracle 主备同步）</Text>
</Box>
```

**带文字标注的横向箭头**：`Box` 高度取 `字号×1.4 + 箭头高 + 2`，一般 **36~40px**（不是 22px，会溢出）。

```
<Box style={{ position:'absolute', top:62, left:136, width:74, height:34,
              flexDirection:'column', alignItems:'center', gap:1 }}>
    <Text style={{ fontSize:13, color:'#1A7FE0', fontWeight:'bold' }}>写入</Text>
    <svg width={74} height={10} viewBox='0 0 74 10'>
        <line x1='0' y1='5' x2='64' y2='5' stroke='#2E7FD4' strokeWidth='2' />
        <polygon points='74,5 62,0 62,10' fill='#2E7FD4' />
    </svg>
</Box>
```

## DSL 硬约束（违一条就报错或渲染异常）

- **禁止** `import` / `export` / `module` / hooks / state / 函数组件定义。文件最后一个表达式必须是 `<Slide>`。
- `Box` **只支持 `display:'flex'`** —— 禁 grid，不支持 `calc()`。
- `box-sizing: border-box`：`width` 已含 padding 和 border。
- Slide 的 `padding` 是安全区，**不能被绝对定位绕过**。
- 换行用 `<br />`，不用 `\n`；局部样式用 `<span style={{...}}>`，不支持 `<strong>`/`<em>`/`<u>`。
- `FAIcon` 的 **`fill` 不设则完全不可见**；图标名必须真实存在，**禁编造**（查 Font Awesome 官方名称）。
- 长文本加 `whiteSpace:'nowrap'` 防非预期折行。
- 禁止自创配色 —— 颜色取自原图。禁止省略原图信息 —— 所有模块名、列表项、图例文字逐字复刻。

### 文案里出现 `{` `}` 必须包成 `{'...'}`

JSX 会把裸花括号当表达式解析。写代码截图、模板占位符这类内容时尤其容易撞上：

```
<Text>{'<div '}</Text>
<Text>{'{/* 省略其他代码 */}'}</Text>
<Text>{'= () => {'}</Text>
```

### 特殊图形用内联 `<svg>`，并可内嵌 `<text>`

需要原图里的自定义插画（等距立方体、线稿占位图、盾牌等）时，`FAIcon` 不够用，直接画 SVG。**`<svg>` 内嵌 `<text>` 可以正常渲染**，中英文都行：

```
<svg width={124} height={92} viewBox='0 0 124 92'>
    <path d='M62 18 L86 31 L62 44 L38 31 Z' fill='#7FB0F7' />
    <text x='62' y='38' fontFamily='Arial' fontSize='13' fontWeight='bold'
          fill='#FFFFFF' textAnchor='middle'>AI</text>
</svg>
```

`overflow:'hidden'` 在 `Box` 上有效，可用于卡片裁剪（让内容不越出圆角）。

## 提交前自检

- [ ] 总高度算过，不溢出 720
- [ ] 每个视觉区块有独立容器 Box，必要时加背景增强分区感
- [ ] 原图所有文字逐字复刻，无省略
- [ ] 配色全部取自原图，未自创
- [ ] 图标语义可辨识
- [ ] lint 通过，**且诊断当视觉检查读过一遍**（见下）
- [ ] 长文本折行后没压到隔壁元素
- [ ] 箭头标签与相邻卡片无交叠

## 关于 lint 的诊断

`child-containment` 与溢出诊断**不只是报错器，是布局瑕疵探测器**。

真实案例：一页架构图上 lint 报了两处溢出，顺着查下去发现两个真实视觉缺陷 ——
1. 一段说明文字实际折成两行，第二行压到了右侧卡片上（文字框宽度给窄了）
2. 路由清单里两个 `<Text>` 被内容撑开，后一个标签被前一个压住

**所以每次 lint 都要当视觉检查读一遍诊断，而不是只跑到 `ok:true` 就结束。** 方框指向哪里，就去追问那里是不是真的挤了。修法：

| 症状 | 修法 |
| :--- | :--- |
| 文字折行压到隔壁 | 缩短文案 / 降字号 / 加宽容器 / 加 `whiteSpace:'nowrap'` |
| flex 行内标签互相重叠 | 给被挤的 `<Text>` 加 `whiteSpace:'nowrap'`，或缩短文案 |
| 箭头标签溢出（`bottom+Npx`） | 高度改成 `字号×1.4 + 箭头高 + 2`（一般 36~40px） |
| 绝对定位元素压在卡片上 | 缩小 width 或左移，给卡片留通道 |

## 视觉验收

本机通常没有 PowerPoint / 完整 WPS，`slidep screenshot` 也可能失败。若需要像素级复核，见 `references/layout-verify.md`。

## 版本并存

矢量版与图片版**都保留**，文件名区分：

- `XXX.pptx` —— 矢量版（要改内容时用）
- `XXX_图片版.pptx` —— 图片版（要绝对保真时用）
