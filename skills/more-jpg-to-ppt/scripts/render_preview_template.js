/**
 * render_preview_template.js —— slidep 页面版式预览图模板
 *
 * 用途：本机没有 PowerPoint / 完整 WPS，且 `slidep screenshot` 失败时，
 *      按 .slide 的绝对定位坐标 1:1 自绘一张预览图做视觉验收。
 *
 * ⚠️ 核心纪律：本脚本的坐标必须与 .slide 严格一一对应。
 *    改完 .slide 一定要同步改本脚本，否则预览图会骗人。
 *
 * 运行：
 *   export NODE_PATH="<node>/node_modules/@tencent/slidep/node_modules"
 *   "<node.exe>" render_preview_template.js
 */

const { Canvas } = require('skia-canvas');

const W = 1280, H = 720;
const canvas = new Canvas(W, H);
const ctx = canvas.getContext('2d');

// ===== 色板：与 DESIGN.md 保持一致，直接填 hex =====
const C = {
  blue:   '#1A7FE0',   // 主蓝：标题条 / 强调
  dark:   '#14456F',   // 深蓝：正文标题
  green:  '#2E9E5B',   // 备库绿 / 正向状态
  red:    '#E8453C',   // 警示红
  body:   '#3D566E',   // 正文灰蓝
  grey:   '#7B8A9A',   // 次要说明
  lblue:  '#EAF3FC',   // 浅蓝底
  lgreen: '#E9F6EE',   // 浅绿底
  lgrey:  '#F4F7FA',   // 浅灰底
  bblue:  '#C9DDF2',   // 边框蓝
  arrow:  '#2E7FD4',   // 箭头蓝
};
const FONT = '"Microsoft YaHei", sans-serif';

// ===== 基础绘制工具 =====

/** 圆角矩形（fill / stroke 都可选） */
function rr(x, y, w, h, r, fill, stroke, lw) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
  if (fill) { ctx.fillStyle = fill; ctx.fill(); }
  if (stroke) { ctx.strokeStyle = stroke; ctx.lineWidth = lw || 1; ctx.stroke(); }
}

/** 单行文字。align 传 'left'|'center'|'right' */
function txt(s, x, y, size, color, weight, align) {
  ctx.font = `${weight || 'normal'} ${size}px ${FONT}`;
  ctx.fillStyle = color;
  ctx.textAlign = align || 'left';
  ctx.textBaseline = 'middle';
  ctx.fillText(s, x, y);
}

/** 自动换行多行文字，返回占用高度 */
function para(s, x, y, size, color, maxW, lh, weight) {
  ctx.font = `${weight || 'normal'} ${size}px ${FONT}`;
  ctx.fillStyle = color;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  let line = '', lines = [];
  for (const ch of s) {
    if (ctx.measureText(line + ch).width > maxW) { lines.push(line); line = ch; }
    else line += ch;
  }
  lines.push(line);
  lines.forEach((l, i) => ctx.fillText(l, x, y + i * lh));
  return lines.length * lh;
}

/** 横向箭头（带三角箭头） */
function arrowH(x, y, len, color, lw) {
  const c = color || C.arrow, w = lw || 2.4;
  ctx.strokeStyle = c; ctx.lineWidth = w;
  ctx.beginPath(); ctx.moveTo(x, y); ctx.lineTo(x + len - 10, y); ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(x + len, y);
  ctx.lineTo(x + len - 12, y - 6);
  ctx.lineTo(x + len - 12, y + 6);
  ctx.closePath(); ctx.fillStyle = c; ctx.fill();
}

/** 纵向虚线双向箭头 */
function arrowVDouble(x, y, h, color) {
  const c = color || C.arrow;
  ctx.save();
  ctx.setLineDash([5, 4]); ctx.strokeStyle = c; ctx.lineWidth = 2;
  const mid = y + h / 2;
  ctx.beginPath(); ctx.moveTo(x, y + 4); ctx.lineTo(x, mid - 6); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(x, mid + 6); ctx.lineTo(x, y + h - 4); ctx.stroke();
  ctx.restore();
  ctx.fillStyle = c;
  ctx.beginPath(); ctx.moveTo(x, y); ctx.lineTo(x - 5, y + 10); ctx.lineTo(x + 5, y + 10); ctx.closePath(); ctx.fill();
  ctx.beginPath(); ctx.moveTo(x, y + h); ctx.lineTo(x - 5, y + h - 10); ctx.lineTo(x + 5, y + h - 10); ctx.closePath(); ctx.fill();
}

// ===== 页面绘制 =====
// 按 .slide 的分区顺序，逐块复刻。下面给一个「带标题条的分区卡」的标准写法。

ctx.fillStyle = '#FFFFFF';
ctx.fillRect(0, 0, W, H);

/**
 * 分区卡 + 标题条
 * @param {number} x @param {number} y @param {number} w @param {number} h
 * @param {string} title 标题条文字
 */
function panel(x, y, w, h, title) {
  rr(x, y, w, h, 8, '#FFFFFF', C.bblue, 1);
  // 标题条需被圆角裁切，才能让四角跟随卡片圆角
  ctx.save();
  rr(x, y, w, h, 8, null, null);
  ctx.clip();
  ctx.fillStyle = C.blue;
  ctx.fillRect(x, y, w, 30);
  ctx.restore();
  txt(title, x + 12, y + 15, 16, '#FFFFFF', 'bold');
}

// —— 示例：中部主区 ——
panel(16, 96, 820, 324, '整体架构');
panel(848, 96, 416, 324, '配置策略');
panel(16, 428, 820, 272, '改造方案');
panel(848, 428, 416, 272, '数据保留与归档');

// —— 示例：深色标签（蓝底白字圆角块） ——
rr(304, 149, 126, 30, 6, C.blue);
txt('主库（写）', 367, 165, 15, '#FFFFFF', 'bold', 'center');

// —— 示例：横向箭头 + 上方标注 ——
txt('写入', 173, 168, 13, C.blue, 'bold', 'center');
arrowH(136, 181, 74, C.arrow, 2);

// —— 示例：纵向虚线双向箭头 + 右侧标签（拆两个 Box，避免交叠） ——
arrowVDouble(473, 188, 66, C.arrow);
txt('数据库同步', 500, 216, 15, C.dark, 'bold');
txt('（如 Oracle 主备同步）', 500, 238, 12, C.grey);

// TODO: 按 .slide 补齐其余分区内容

// ===== 输出 =====
// ⚠️ 用 toFile，不要用 saveAs（已废弃且返回 undefined）
canvas.toFile('D:/WorkBuddy/0.任务/_preview_render.png');
console.log('rendered: _preview_render.png');
