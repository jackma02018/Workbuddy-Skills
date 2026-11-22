# 环境约束与避坑（本机 Windows 实操记录）

> 这些坑踩过一次，写下来避免重复。执行前扫一眼，能省 20 分钟。

## 一、必须用绝对路径调 Python

本机 **Bash 的 PATH 缺 coreutils**（`dirname`、`cd`+相对路径、`ls`、`cp`、`head` 可能都不可用），
且 `python` / `pip` 裸命令不一定在 PATH 上。统一用：

```
PY = C:/Users/majunpeng/.workbuddy/binaries/python/envs/default/Scripts/python.exe
```

已装的库：`python-pptx`、`Pillow`。需要新库时装进同一个 venv，**不要 `pip install -g`**。

## 二、工具使用避坑

| 现象 | 原因 | 对策 |
|---|---|---|
| `ls` / `cp` / `head` 报 command not found | Bash PATH 缺 coreutils | 用 Python 的 `os.listdir` / `shutil.copyfile` 替代 |
| PowerShell 命令执行成功但看不到输出 | 本机 PowerShell 工具不回传 stdout | 先写文件再用 Read 读 |
| Read / Grep 读工作区外路径被拒 | 工具限工作区 | 把中间产物写到工作区内再读 |
| 写 `.ps1` / `.bat` 处理中文路径出错 | 脚本文件编码 | 中文路径操作一律走 Python，不写批次脚本 |
| 生成脚本里中文注释乱码 | 编码 | 源码首行 `# -*- coding: utf-8 -*-`，写文件用 `io.open(..., encoding='utf-8')` |

## 三、PPT 相关

- **本机没有 PowerPoint / WPS**：无法做 Office 端像素验收。交付时如实说明
  "版式已通过字体度量自检，未做 Office 渲染验收"，不谎称已肉眼验证。
- **版式正确性靠 `check_layout.py`**：基于 `C:/Windows/Fonts/msyh.ttc` 真实度量，
  检测文本溢出文本框 / 越出卡片 / 越出页面。**交付前必须跑，要求 0 问题。**
- **python-pptx 的两个坑**：
  1. `add_textbox` 默认带一个空段落 → 复用时优先用 `paragraphs[0]`，避免多出空行；
  2. 需要重新生成整页时直接重跑脚本，不要"在已有 pptx 上打补丁"——补丁式修改易出现
     形状残留与坐标漂移。
- **图片复制走原始字节**：`Part` + `relate_to`；**禁用** `get_or_add_image_part`
  （依赖 PIL，EMF/WMF/SVG 会破图）——多页合并、素材套用场景必守。
- **矢量优先于位图**：本套方案全部用原生形状画图标，好处是无外部素材依赖、可编辑、可改色、
  打印不失真；不要为了"看起来专业"去贴网图（版权与清晰度双重风险）。

## 四、素材获取

- **ima 知识库只返回解析后的文本**，没有图片 / 文件二进制下载通道。
  需要"素材图"时，改为按语义矢量重绘（本技能的做法），不要指望从 ima 取出原图。
- 检索 ima 用 `get_knowledge_list` + `folder_id` 翻页；**大范围查找不要用
  `search_knowledge`**——结果会被截断落盘到工作区外读不到的临时路径。
- 需要往 ima 存产出时，走 `创建媒体 → COS 签名上传 → 登记知识库` 三步，
  可复用技能 `ima-knowledge-archive`（含上传脚本）。

## 五、命名与归档约定

- 交付件命名：`客户名_主题_版本.pptx`（例：`绿控传动_生产管理专题方案_5页.pptx`）。
- 中间产物（内容 JSON、脚本）与交付件同目录，便于复现：
  `build_xxx_ppt.py` + `check_xxx.py`，或本技能的 `content_*.json`。
- 入库知识库时统一命名 `主题-日期-版本`（例：`MES方案-绿控传动生产管理专题PPT-20260917-v3`）。
