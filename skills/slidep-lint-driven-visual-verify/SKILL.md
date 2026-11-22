---
name: slidep-lint-driven-visual-verify
description: >-
  在无 PowerPoint/WPS 的机器上校验 slidep 生成的 .pptx 版式：用 lint 的 child-containment 诊断当视觉检查器，
  并用 @tencent/slidep 自带的 skia-canvas 按绝对定位几何 1:1 绘制预览图做像素级复核。
  当 slidep screenshot 报 `file:///workspace/...` 路径错误、或需要脱离 Office 验证 PPT 版式时使用。
agent_created: true
---

# slidep 版式校验（无 Office 环境）

## 适用场景

- `slidep screenshot` 失败，报 `render failed: The argument 'filename' must be a file URL object...Received 'file:///workspace/node_modules/...'`
- 本机没有 Microsoft PowerPoint，也没有完整 WPS（只有 `wpscloudsvr.exe` / Kingsoft PDF）
- 需要确认 `slidep` 生成的 `.pptx` 里文字有没有重叠、溢出、被裁切

## 核心认知

**`slidep lint` 的 `child-containment` 与溢出诊断不只是"报错器"，它是布局瑕疵探测器。**

真实案例：一页 1280×720 的架构图上，lint 报出两处 `child-containment` 溢出，顺着查下去发现的是两个真实视觉缺陷：
1. 一段说明文字实际折成了两行，第二行压到了右侧卡片上（文字框宽度给窄了）
2. 路由清单里两个 `<Text>` 的 flex 分配被内容撑开，路由标签"走读库"被前一个标签压住

**所以：每次 lint 都要当视觉检查读一遍诊断，而不是只把它跑到 `ok:true` 就结束。** 溢出方框指向哪里，就去追问那里是不是真的挤了。

## 步骤

### 1. 先跑 lint，逐条读诊断

```bash
slidep lint "<file>.pptx" --dsl-file "slides/01.slide"
```

返回 `diagnostics` 数组，每条含 `ruleId` / `severity` / `message`。`message` 里会打印出出问题的 JSX 片段和溢出方向（`bottom+14px` 等）。**逐条去看那段 JSX，判断是"改小字号"还是"挪位置"还是"放宽容器"。**

### 2. 用 skia-canvas 自绘预览图

`@tencent/slidep` 依赖树里带着 `skia-canvas`，可直接用：

```bash
export NODE_PATH="C:/Users/<user>/.workbuddy/binaries/node/versions/<ver>/node_modules/@tencent/slidep/node_modules"
"<node.exe>" render_preview.js
```

渲染脚本要点：
- `const { Canvas } = require('skia-canvas')`
- 画布就用 slide 的真实尺寸（1280×720），**不要把 DSL 的 px 值换算**，直接按同一套绝对定位坐标画
- **用 `canvas.toFile(path)`，不要用 `canvas.saveAs()`**（后者已废弃且返回 `undefined`，`.then()` 会 TypeError）
- 圆角矩形用 `arcTo` 手写；中文用 `ctx.font = 'bold 15px "Microsoft YaHei", sans-serif'`
- 文字 `textBaseline = 'middle'` 时 y 坐标取"文字行垂直中心"，与 CSS flex 居中更接近

**关键纪律：预览脚本必须与 `.slide` 的坐标严格一一对应。改完 `.slide` 一定要同步改渲染脚本，否则预览图会骗人。**

### 3. 看图，找重叠 / 溢出 / 裁切

重点扫三处：
- 长文字有没有折行后压到隔壁元素（DSL 里算得下，渲染出来不一定）
- 箭头标签与相邻卡片有没有交叠
- 底部内容有没有贴近容器底边

### 4. 修 `.slide` → 重 lint → 重新 upsert

```bash
slidep upsert-dsl "<file>.pptx" --dsl-file "slides/01.slide" --page-index 0
```

## 常用修法

| 症状 | 修法 |
| :--- | :--- |
| 文字折行压到隔壁 | 缩短文案 / 降低字号 / 加宽容器 / 加 `whiteSpace: 'nowrap'` |
| flex 行内标签互相重叠 | 给被挤的 `<Text>` 加 `whiteSpace: 'nowrap'`，或缩短文案 |
| 箭头标签溢出（`bottom+Npx`） | 标签 Box 高度 = `字号×1.4 + 箭头高度 + 2`，一般 36~40px；或拆成"标签 Box + SVG Box"两个绝对定位 Box |
| 绝对定位元素压在卡片上 | 缩小 width 或左移，给卡片留出通道 |

## 环境备注

- slidep 在 Node 托管环境里：`<node>/node_modules/@tencent/slidep`，可执行文件 `<node>/slidep`
- bash 缺 PATH 时先 `export PATH="/usr/bin:/bin:/c/Windows/System32:/c/Windows:$PATH"`
- 不要用 `reg.exe` 探测 Office 安装（常被沙箱黑名单拦截），直接 `find` 找 `POWERPNT.EXE` / `wpp.exe`
