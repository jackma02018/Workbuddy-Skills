# 技能包备份到资料库（含 doc 编辑接口踩坑）

> 为什么需要：技能装在 `~/.workbuddy/skills/` 是**纯本地**的，重装系统即清空。
> 资料库里的文档只是**参考材料**，不会被自动调用（没有 `SKILL.md` 结构），所以必须备份一份**可还原的包**。

## 1. 打包

目录结构（外层文件夹名 = 技能名）：

```
more-jpg-to-ppt/
├── SKILL.md
├── 恢复说明.md            # 离线版恢复步骤，解压即可看
├── references/*.md
└── scripts/*
```

打包（用 Python 标准库，跨平台无依赖）：

```python
import zipfile, os
src = "<pkg_dir>"          # 含 more-jpg-to-ppt/ 的父目录
out = "<out>.zip"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for root, _, files in os.walk(src):
        for f in files:
            p = os.path.join(root, f)
            z.write(p, os.path.relpath(p, src))
print(zipfile.ZipFile(out).testzip())   # 必须为 None
```

`testzip()` 返回 `None` 才算完整。

## 2. `.zip` 的三条路由（关键结论）

资料库有多个容器，但**只有附件能装 `.zip`**：

| 容器 | 结果 | 原因 |
| :--- | :--- | :--- |
| `drive` 网盘 | ❌ `该文件类型不走网盘上传，请使用对应资料库模块` | 后端 `uploadType` 把 `.zip` 判给 page，本地脚本直接拒绝 |
| `page` 页面 | ❌ 不适用 | 契约要求 zip 内至少含一个 `.html` / `.htm` 入口 |
| `attachment` 节点附件 | ✅ | 无扩展名或类型限制 |

不要试图给 `.zip` 里塞一个空 HTML 来骗过 page 契约——契约明确写着「无入口时先和用户确认，**不要打空入口 zip**」。

## 3. 上传为附件

```bash
python "<SKILL_DIR>/attachment/attachment.py" upload --token-stdin \
  --node-id <所属文档的 nodeId> \
  "<技能包路径>.zip" \
  --file-name "技能包_YYYYMMDD.zip"
```

- `--node-id` 必须是已存在的 **doc / page / database** 节点，附件跟着它走、不在目录树单独占位。
- 成功输出 `KS_ATTACHMENT_UPLOAD_OK <json>`，含 `attachment_id`（**要记下来**，下载时必传）。
- 回执原样透传 `KS_USER_REPLY`。

验收（只读）：

```bash
python "<SKILL_DIR>/attachment/attachment.py" download --token-stdin \
  --node-id <nodeId> --attachment-id <att_id>
```

能拿到 `KS_ATTACHMENT_DOWNLOAD` 即通。`download_url` 是短期预签名链接，不要整串回显给用户。

## 4. 在文档正文写恢复说明

用 `doc/get_doc_reviews.py` 回读拿最新块，再 `doc/submit_doc_edit.py` 追加。
**文末追加**用 `insert_after` + `id:""`。

### 4.1 踩坑清单（逐个都撞过）

1. **`--actions-file` 必须传 Windows 路径**。传 `/tmp/xxx.json` 会报 `读取 actions-file 失败`。
   Git Bash 下先转换：`cygpath -w /tmp/actions.json` → `C:\Users\...\Temp\actions.json`。
2. **`content` 必须是组件语法，禁止裸 Markdown**。`` `code` ``、`**bold**`、`- 列表` 全部会被拒：
   `content/new_content 必须使用组件语法，禁止裸 Markdown 或纯文本`。
3. **`<Mark>` 只有这些属性**：`ar, backgroundcolor, bold, color, comment, italic, strike, underline`。
   **没有 `code` 属性** —— 行内代码样式无法表达，只能改写成普通文字加引号或书名号。
4. **`<Table>` 不接受任何属性**。`readonly` / `rowHeader` 会被拒（`仅允许: 无属性`）。
5. `<TableCell>` 内文字必须由 `<Paragraph>` 承载，禁止裸写。
6. 直接编辑模式下 `content` 禁 `<Mark ar>`；只有审阅模式才用 `ar`。

### 4.2 可用的组件骨架

```xml
<Heading level="2">九、备份与恢复</Heading>

<BlockQuote><Paragraph>保存日期 YYYY-MM-DD</Paragraph></BlockQuote>

<Paragraph>正文，行内样式用 <Mark bold>加粗</Mark>。</Paragraph>

<Code language="bash">pip install python-pptx pillow</Code>

<Table>
  <TableRow>
    <TableCell><Paragraph>表头</Paragraph></TableCell>
  </TableRow>
</Table>
```

块与块之间用**空行**分隔；`<Code>` 内容按纯文本处理，内部不写 `<Paragraph>` / `<Mark>`。

### 4.3 提交与验收

```bash
printf '%s' "<token>" | python "<SKILL_DIR>/doc/submit_doc_edit.py" --token-stdin \
  --page-id <nodeId> --actions-file "C:\...\actions.json"
```

成功输出 `KS_DOC_EDIT_SUBMIT` + `affectedBlockIds`，并原样透传 `KS_USER_REPLY`。
回读复核用 `doc/get_doc_reviews.py --page-id <nodeId>`。

> `nodeId` 与 `pageId` 在 Doc 接口里等价，统一传 `--page-id`。

## 5. 恢复说明该写什么（五步）

正文里必须交代清楚，尤其是**第 2 步的常见错误**：

1. **下载**：在文档附件区找到技能包 zip，解压。
2. **放回路径**：整个技能文件夹放到 `C:\Users\<用户名>\.workbuddy\skills\`，正确结构是 `skills\more-jpg-to-ppt\SKILL.md`。
   ⚠️ **不要多套一层同名目录**（变成 `skills\more-jpg-to-ppt\more-jpg-to-ppt\SKILL.md`），否则技能不被识别。
3. **装依赖**：`pip install python-pptx pillow`。矢量版另需 slidep（Node 环境）与 skia-canvas（随 slidep 自带）。
4. **重启** WorkBuddy。
5. **验证**：发一张图说「转成 PPT」，能触发即成功。
