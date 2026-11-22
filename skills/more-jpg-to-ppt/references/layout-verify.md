# 无 Office 环境下的版式验收

本机通常**没有 Microsoft PowerPoint，也没有完整 WPS**（只有 `wpscloudsvr.exe` / Kingsoft PDF），无法用 Office 导出图片验收。`slidep screenshot` 也可能报错：

```
render failed: The argument 'filename' must be a file URL object, file URL string,
or absolute path string. Received 'file:///workspace/node_modules/@tencent+docs-slide-view-sdk-node@...'
```

不必卡在这里。下面两条路可以先做。

## 路一：认真读 lint 的诊断

见 `vector-mode.md` 的"关于 lint 的诊断"一节。**这是成本最低、命中率最高的一道关**，能把大部分真实版式缺陷挡在前面。

## 路二：用 skia-canvas 自绘预览图

`@tencent/slidep` 的依赖树里带着 `skia-canvas`，可以直接拿来按 DSL 的绝对定位坐标 1:1 重画一张预览图。

```bash
export NODE_PATH="C:/Users/<user>/.workbuddy/binaries/node/versions/<ver>/node_modules/@tencent/slidep/node_modules"
"C:/Users/<user>/.workbuddy/binaries/node/versions/<ver>/node.exe" render_preview.js
```

### 脚本要点

- `const { Canvas } = require('skia-canvas')`
- 画布就用 slide 的真实尺寸（1280×720）。**不要把 DSL 的 px 值做任何换算**，直接按同一套绝对定位坐标画 —— 这样预览图才和 PPT 一致。
- **用 `canvas.toFile(path)`，不要用 `canvas.saveAs()`**：后者已废弃，且返回 `undefined`，接 `.then()` 会 TypeError。
- 圆角矩形用 `arcTo` 手写：
  ```js
  ctx.beginPath();
  ctx.moveTo(x+r, y);
  ctx.arcTo(x+w, y, x+w, y+h, r);
  ctx.arcTo(x+w, y+h, x, y+h, r);
  ctx.arcTo(x, y+h, x, y, r);
  ctx.arcTo(x, y, x+w, y, r);
  ctx.closePath();
  ```
- 中文：`ctx.font = 'bold 15px "Microsoft YaHei", sans-serif'`
- 文字 `textBaseline = 'middle'` 时，y 取"文字行垂直中心"，更接近 CSS flex 居中。
- `textAlign` 用 `'left'|'center'|'right'`。

### 核心纪律

**预览脚本必须与 `.slide` 的坐标严格一一对应。改完 `.slide` 一定要同步改渲染脚本，否则预览图会骗人 —— 这比没有预览更危险。**

#### 七个真实的翻车点（务必规避）

1. **改了 DSL 忘了改脚本** —— 渲染出来跟没改一样，会误判成"改动无效"而反复改 DSL。
   规避：把"改 `.slide`"和"改 preview 脚本"当成一个动作，成对执行。
2. **同名文件被宿主图片缓存命中** —— 重新渲染到同一个文件名，文件 hash 不变时宿主会判定"图片未更新"，你看到的还是旧图。
   规避：**每次渲染换个新文件名**（`_v2.png` / `_v3.png`），用完删掉。
3. **忘了加卡头高度（最容易漏，且症状极具迷惑性）** ——
   带标题条的分区卡，DSL 里 `overflow:'hidden'` 的内容区是 `<Box style={{height:44}}>` + `<Box style={{position:'relative', flex:1}}>` 两个 flex 子项。
   分区内绝对定位元素的 `top` 是**相对那个 `flex:1` 容器**的，所以**绝对 y = 卡 top + 卡头高 + 内容 top**。
   渲染脚本里如果直接用 `卡 top + 内容 top`，**所有内容都会上移一个卡头的高度**，症状是"卡头文字与正文/标题重叠"。
   → 你会误判成"DSL 里正文 top 给小了"，然后去改 DSL，越改越乱。
   **规避：渲染脚本里把卡头高度显式加进每个 y 坐标，或给每个分区设一个 `const BODY = 卡top + 44;` 基准常量。**

```js
// 正确姿势：显式加卡头
const BODY1 = 80 + 44;               // 分区 1 卡体原点
txt('正文首行', 30, BODY1 + 20, ...);  // 而不是直接写 80 + 20
```

4. **skia-canvas 的导入方式是 `{ Canvas }` 类，不是 `createCanvas` 工厂函数** ——
   ```js
   // 错：TypeError: createCanvas is not a function
   const { createCanvas } = require('skia-canvas');
   const canvas = createCanvas(1280, 720);

   // 对
   const { Canvas } = require('skia-canvas');
   const canvas = new Canvas(1280, 720);
   ```
   该模块导出的是 `Canvas` / `CanvasRenderingContext2D` / `Image` / `Path2D` 等**类**，没有 `createCanvas`。
   排查手法：`node -e "console.log(Object.keys(require('.../skia-canvas')))"` 直接把导出键打出来看。

5. **`ctx.ellipse` 画圆柱体分层时，先画体再画顶面椭圆** ——
   数据可视化常见的"数据库圆柱堆叠"图标：每层的侧壁是 `moveTo + lineTo + ellipse(下半弧) + lineTo + closePath`，
   然后**单独补一个完整椭圆做顶面**。顺序反了会被侧壁盖住，看起来像实心柱子。

6. **改"通用绘制函数"会连带改坏所有调用点（本轮新增，代价 2 轮渲染）** ——
   `preview.js` 里 `robot(x, y, size, body, face)` 被 6 处复用：①的头像、②的中央大图标、③笔记本里的小头像、
   ③手机里的小头像、底部品牌位。我给其中一处调机身几何，结果**① 原本的"圆形用户头像"被一起改成了机器人形状**
   （渲染出来像块棺材板），另外几处的耳翼/天线也全变了。
   **规避：改函数签名或内部路径前，先 `grep -n "函数名("` 把所有调用点列出来。**
   同一张图里"看起来同类"的角色，原图往往是**不同造型**，应该拆成两个独立函数
   （本次拆成 `robot()` 与 `userAvatar()`），而不是用一个通用函数 + 参数硬凑。

7. **同一个角色在不同位置，原图细节可能完全不同（本轮新增）** ——
   本图里 WisCode 机器人出现了 4 次，逐处裁开放大后是这样的：

   | 位置 | 造型 |
   | :--- | :--- |
   | ① 对话头像 | 浅蓝圆底 + **双层机身**（深蓝外圈 + 中蓝内芯）+ 左右耳翼 + 天线圆点 + **大椭圆白眼** + 白色嘴缝 |
   | ② 中央大图标 | 同上，但尺寸大、耳翼更明显 |
   | 品牌位底部 | 机身色阶**反过来**：浅蓝在上、深蓝在下 |
   | ③ 设备内小头像 | 简化版，细节不可辨 |

   **规避：不要"看一次推广到全部"。每个位置单独裁开放大确认，或者明确接受"小尺寸处可简化"并记在 DESIGN.md 里。**

### 看图时重点扫三处

1. 长文字有没有**折行后压到隔壁元素**（DSL 里算得下，渲染出来不一定）
2. 箭头标签与相邻卡片有没有**交叠**
3. **容器内是否留有大片空白** —— 常见于"子元素依次排下来只占了 470/504 高"，此时应把页脚/收尾元素改为贴底绝对定位，或放大内容间距把空间吃满，别让卡片底部空着

### 图解元素要回原图对照，别凭印象画

本轮的教训：图标画完觉得"差不多"，放大原图一看完全不是一回事 ——
页头徽标原图是**蓝色实心折角图纸 + 淡蓝左侧竖条 + 四角取景框**，我第一版画成了线框；
流程条 5 个图标原图全是**实心蓝阶图形**（文档 / AI 芯片 / 数据库堆叠 / 人群 / 漩涡），我第一版画成了线框描边 + 实心方块占位。

**规矩：凡是图标、徽标、示意图这类"造型"元素，动手前先把原图对应区域裁出来放大看**
（`PIL` 裁切 + `resize` 放大 3 倍），确认是实心还是线框、有几层色阶、外轮廓什么形状，再写 SVG。
这类偏差 lint 查不出来（几何完全合法），只有像素级回看才会暴露，而它恰恰最影响"像不像"。

### 完整示例

见 `scripts/render_preview_template.js`。把它复制到项目目录，按 `.slide` 的分区逐个改坐标和文案。

## 兜底

以上都不适用时，直接把 `.pptx` 交给用户，并**明确说明未能做像素级验收**，请其在 PowerPoint/WPS 里过一眼。不要假装验过了。
