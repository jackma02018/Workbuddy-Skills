# OOXML / python-pptx 深水区笔记

本文件是本技能踩坑后的沉淀，遇到"改完报错 / 图片破损 / 组无效"先查这里。

## 0. 运行环境（本机）

- python-pptx 1.0.2，解释器固定用绝对路径：
  `C:\Users\majunpeng\.workbuddy\binaries\python\versions\3.13.12\python.exe`
- **bash 被 shim 破坏**：`dirname` / `cd` / `ls` / `head` 均报 `command not found`，且 stderr 会混进输出。
  因此：不要依赖 shell 内置命令，直接用绝对路径调 Python 执行脚本；文件列举/读取用 Python 或专用工具。
- **本机无 PowerPoint / WPS**：任何"看起来怎样"的结论都无法本地验收，必须如实说明。

## 1. 图片：必须走原始字节复制

python-pptx 的 `SlidePart.get_or_add_image_part(image_file)` 内部用 PIL 识别图片，
遇到 **EMF / WMF / SVG** 会抛 `cannot identify image file`，结果是图片破损或异常。

正确做法（复制幻灯片时）：

```python
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

part = Part(PackURI("/ppt/media/image%d%s" % (n, ext)), src_part.content_type, package, src_part.blob)
new_rId = dst_slide_part.relate_to(part, rel.reltype)     # rel.reltype 通常是 RT.IMAGE
```

- `Part(partname, content_type, package, blob)` —— 第四个参数是原始字节，不经过 PIL。
- `relate_to(target, reltype, is_external=False)` —— 外部链接（超链接）用 `is_external=True` + URL 字符串。
- 全文档按 `(reltype, sha1(blob))` 去重，避免同一 Logo / 图标被复制 N 份。
- 若源包只把图片放在 `ppt/slides/media/`（非标准的 `ppt/media/`），目标端统一写到 `/ppt/media/` 即可。

## 2. 复制幻灯片后必须重映射 rId

深拷贝 `spTree` 只是搬 XML；形状上的 `r:embed` / `r:link` / `r:id` 仍指向**源包**的关系 ID。
必须遍历新页所有元素，把 `{...relationships}` 命名空间下的每个属性值换成新包中的 rId：

```python
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
for el in dst_tree.iter():
    for attr in [a for a in el.attrib if a.startswith(NS_R)]:
        rid = el.get(attr)
        rel = src_part.rels[rid]
        ...  # 建新关系 -> new_rId
        el.set(attr, new_rId)
```

`slideLayout` / `slideMaster` 关系不会出现在形状属性里，不受影响。

## 3. 字体：中文字体要写三处，字号是属性

- 只写 `a:latin` 时中文仍回退到主题字体 → 必须同时写 `a:latin` / `a:ea` / `a:cs` 的 `typeface`（如 `黑体`）。
- 子元素顺序有约束：`latin/ea/cs` 要排在 `sym` / `hlinkClick` / `hlinkMouseOver` / `rtl` / `extLst` **之前**，直接 `append` 可能产生非法顺序。
- 字号写在 `a:rPr` / `a:defRPr` / `a:endParaRPr` 的 **`sz` 属性**上（不是子元素），单位百分之一磅：`1200` = 12pt。
- `a:rPr` 没有 `sz` 表示继承，别当成 0 处理；判定"标题"时按有 `sz` 的 run 取最大值。

### 陷阱：`font.size = 浮点数` 会报错

```python
run.font.size = 11.2              # ✗ ValueError: value must be in range 100 to 400000 inclusive, got 0
run.font.size = Pt(11.2)          # ✓ 必须传 Length（Pt/Emu）
```

原因是 setter 期望 Length（EMU），裸浮点被当成 EMU 再换算成百分之一磅后归零。

## 4. 组（p:grpSp）与等比缩放

结构：

```xml
<p:grpSp>
  <p:nvGrpSpPr>...</p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="274320" y="731520"/>      <!-- 屏幕上组的左上角 -->
      <a:ext cx="10363200" cy="5829300"/> <!-- 屏幕上组的尺寸 = scale × chExt -->
      <a:chOff x="0" y="0"/>              <!-- 子坐标系的左上角 -->
      <a:chExt cx="12192000" cy="6858000"/><!-- 子坐标系尺寸 = 原始整页 -->
    </a:xfrm>
  </p:grpSpPr>
  ...子形状（坐标仍是原始绝对坐标）...
</p:grpSp>
```

- **等比缩放的本质就是 `ext = scale × chExt`**，`chOff=(0,0)`、`chExt=整页尺寸` 时子形状可直接沿用原绝对坐标。
- **`a:xfrm` 在 `p:grpSpPr` 里面**，不是 `p:grpSp` 的直接子节点：
  `grp.find(".//" + NS_A + "xfrm")` 才能取到，直接 `find(NS_A+"xfrm")` 返回 `None`。
- 用 `slide.shapes.add_group_shape([shape, ...])` 建组时，传入的形状会被**移动**（lxml 语义）进组，顺序保持，并自动 `recalculate_extents()`；随后覆盖 `a:xfrm` 即可。
- 组缩放会**连带缩放组内文字**，所以组内文字与容器同比例，不会出现"字变大了框没变"的失配。这也是为什么不要对组内文字做批量缩字号。
- z 序 = `spTree` 文档顺序，**末尾在最上层**。把标题/Logo 从树中 `remove` 再 `append`，即可保证不被组遮挡。

## 5. 标题与右侧 Logo 的识别

- 标题：顶层 `p:sp`，有文本，且有显式字号 `≥ 20pt` 且宽度 `> 0.4in`；全都不满足时回退取"最靠上且宽度 > 1.5in 的文本"。
- 右侧公司图标：顶层 `p:pic`，且 `left + width > 0.88 × 页宽`。
- 有的封面页 Logo 来自**版式/母版**（页面上没有 `pic`），此时该页无 Logo 可保留，属正常。

## 6. 文字自适应（normAutofit）

- 位置：`a:bodyPr` 内的 `a:normAutofit`，要排在 `a:prstTxWarp` 之后、`a:scene3d` / `a:sp3d` / `a:flatTx` / `a:extLst` 之前。
- 想"现在就到正确大小"，就得显式改 `sz`；`normAutofit` 负责的是**以后编辑时自动缩小**。
- `bodyPr` 的 `wrap="none"` 表示不换行，溢出表现为**横向**；配合默认内边距（左右 0.1in、上下 0.05in）容易误判为溢出。
  典型假溢出：无边框小标签框（高 0.15in 装 12pt 字），原始海报里本就贴合。
  **因此不要对组内正文批量缩字号**，只处理顶层标题区文字。

### 精准定位法（确需处理组内某个文本框时）

1. 在 `spTree` 中按文本内容找到该 `p:sp`；
2. 判断它属于哪一层（在哪个 `grpSp` 内、组缩放比是多少）；
3. 在**该文本框自身坐标系**内做溢出判断与改字号——由于组内文字与容器同比例缩放，两者都以同一系数变化，可直接在子坐标系内比较；
4. 改完用 `verify_pptx.py` 复验。

## 7. 空白页复刻封面版式

页面上看不到绿点/弧线/Logo 等装饰时，它们来自**版式**。把空白页的版式关系重定向到封面所用版式：

```python
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

old = next(rId for rId, rel in blank.part.rels.items() if rel.reltype == RT.SLIDE_LAYOUT)
del blank.part.rels._rels[old]                      # 先摘掉旧关系，避免同类型两条
blank.part.relate_to(cover_layout_part, RT.SLIDE_LAYOUT)
```

再深拷贝封面页的标题形状到该页（占位符会按新版的版式继承），最后交给 `fit_text.py` 做自适应。

## 8. 校验时的两个"假异常"

1. **zip 目录占位项**：WPS 等工具存盘时会在包里写入 `ppt/slides/media/`（以 `/` 结尾）这类目录条目，
   而 python-pptx 存盘会剔除它们。比对媒体文件时**必须排除以 `/` 结尾的名字**，否则会误报"缺失 3 张图片"。
2. **python-pptx 会剔除未引用的 Part**：重存后包里以 `/` 结尾的目录条目消失是正常现象；
   真实图片数量与 sha1 才是判断依据。

## 9. 其它

- 删除模版多余的页：`sldIdLst.remove(sldId)` 前先 `prs.part.drop_rel(rId)`，否则残留孤儿关系。
- `Relationships` 可用 `.rels[rId]` 取值、`items()` 遍历、`_rels` 字典用于删除。
- 每次改动前备份，命名 `原名_<阶段>前备份.pptx`，便于回溯与 sha1 比对。
