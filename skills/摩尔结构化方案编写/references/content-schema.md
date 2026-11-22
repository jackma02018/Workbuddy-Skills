# 内容 JSON 规范（build_deck.py 输入）

> 引擎与内容解耦：换客户、换行业、换模块只改这份 JSON，`deck_kit.py` 一行不动。
> 生成模板：`<PY> scripts/build_deck.py --dump-example`

## 一、顶层结构

```json
{
  "meta": {
    "corner": "客户名 · MES 建设方案",
    "source": "数据来源：现场调研记录 ｜ 指标为行业基准示意口径，正式立项以现场调研数据替换",
    "brand": "摩尔元数 · MES 方案",
    "out": "E:/xx/客户名_生产管理专题方案.pptx"
  },
  "pages": [ { /* 第 1 页 */ }, { /* 第 2 页 */ } ]
}
```

`meta` 全部可选；`total`（页眉 "01 / 05" 的分母）由页数自动推导，不用手填。
`out` 为相对路径时按当前工作目录解析；命令行 `--out` 优先级更高。

## 二、单页字段

```json
{
  "bone": "A",
  "title": "构建工厂级统一生产管理体系",
  "sub": "以统一工单模型与执行标准，承接 ERP 计划、驱动车间执行",
  "need": ["统一管理语言：8 条线体执行标准一致", "..."],
  "pain": [["系统烟囱林立", "园区自研、吴江自研、外购 3 套并存"], "..."],
  "flow": {
    "name": "流程型思维",
    "sub": "端到端执行主流程 · 六步闭环",
    "steps": [["工单接入", "ERP 工单自动同步"], "..."],
    "note": "关键卡控点：拆解规则 · 齐套校验 · 开工校验 · 完工校验"
  },
  "alt": {
    "name": "系统思维",
    "sub": "三层架构 · ISA-95",
    "color": "5B5FC6",
    "cards": [["factory", "L3 工厂级 MOM", "统一模型 / 编码 / 工艺路线库"], "..."],
  },
  "flow_icons": ["doc", "split", "check", "people", "scan", "arrow"],
  "kpis": [["3 → 1", "系统套数收敛"], "..."],
  "claim": "以一套工厂级 MOM 统一替代 3 套产线级 MES，实现执行标准、数据口径与运维体系的全面收敛。"
}
```

### 字段约束（校验失败会直接报错，提示第几页缺什么）

| 字段 | 必填 | 数量 / 长度 | 说明 |
|---|---|---|---|
| `bone` | 否 | `A`–`E` | 缺省按 `A,B,C,D,E` 顺序轮转 |
| `title` | 是 | ≤ 22 字 | **结论句**，不写"XX分析" |
| `sub` | 是 | ≤ 34 字 | 一句话说清"承接什么、驱动什么" |
| `need` | 是 | 4 条，每条 ≤ 20 字 | 动词开头，业务结果导向 |
| `pain` | 是 | 4 条，`[短标签, 描述]` | 标签 ≤ 8 字；描述 ≤ 22 字且含数字 |
| `flow` | 是 | `name/sub/steps/note` | steps 5–6 步，每步名 ≤ 6 字、说明 ≤ 12 字；note ≤ 34 字 |
| `alt` | 是 | `name/sub/color/cards` | cards 4–6 张；卡标题 ≤ 12 字、说明 ≤ 16 字 |
| `flow_icons` | 否 | 与 `steps` 等长 | 不等长时自动用 `check` 补齐 / 截断；key 见 icon-library.md |
| `kpis` | 是 | 3 或 4 项 | `[数值, 指标名]`；数值 ≤ 8 字，指标名 ≤ 10 字 |
| `claim` | 否 | ≤ 70 字 | 结论收口，落在经营收益 |

### 兼容写法
- `alt.cards` 支持两种：`["gear", "标题", "说明"]` 或 `["gear", "标题"]`（第二视角为纯标签宫格时用，如骨架 C）。
- 颜色可写 `"5B5FC6"` / `"#5B5FC6"` / `[91,95,198]`；不写则用默认色。
- 数组与元组等价（JSON 里都是数组），引擎内部会规范化。

## 三、超出字数怎么办

版面是定死的卡片，**文字超了不要改引擎，按下面顺序处理**：

1. **砍字**：把"实现生产过程的全面数字化管理"压成"生产过程数字化"；
2. **换术语**：长句换行业通用简称（"有限产能排产"而不是"基于设备产能约束的排产计算"）；
3. **降层级**：说明性内容从卡片挪进 `note` 或 `claim`；
4. 最后才考虑调骨架的 `cell_h` / 字号（改引擎参数需重新体检全篇）。

## 四、命令速查

```bash
PY="C:/Users/majunpeng/.workbuddy/binaries/python/envs/default/Scripts/python.exe"

# 生成（自动跑版式体检）
"$PY" scripts/build_deck.py content.json --out "E:/xx/客户_专题方案.pptx"

# 只生成不体检
"$PY" scripts/build_deck.py content.json --no-check

# 单独体检
"$PY" scripts/check_layout.py "E:/xx/客户_专题方案.pptx"

# 导出内容模板
"$PY" scripts/build_deck.py --dump-example
```

## 五、依赖

`python-pptx`、`Pillow`（体检用）。安装（本机隔离 venv，勿污染全局）：

```bash
"C:/Users/majunpeng/.workbuddy/binaries/python/envs/default/Scripts/python.exe" -m venv "C:/Users/majunpeng/.workbuddy/binaries/python/envs/default"
"C:/Users/majunpeng/.workbuddy/binaries/python/envs/default/Scripts/pip.exe" install python-pptx Pillow
```

本机该 venv 已装好，通常无需重装。
