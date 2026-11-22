# 矢量图标库语义索引（26 key）

> 全部由 `deck_kit.glyph()` 用原生矩形 / 椭圆 / 弧线拼绘，**零位图依赖**，可无损放大、
> 可改色、可拆分编辑。选图标先查本表，**按语义选，不按形状像不像选**。
> 用法：`flow_icons: ["doc", "split", ...]` 或 `alt.cards: [["gear", "标题", "说明"], ...]`
> 未识别的 key 会退化为一个圆形兜底图元（不报错，但请避免）。

## 一、主表

| key | 语义 | 典型使用场景 | 常搭配 |
|---|---|---|---|
| `doc` | 工单 / 单据 | ERP 工单接入、计划接收、单据流转 | split、check |
| `split` | 工单拆解 | 工单拆至产线/工序级、任务分解 | doc、people |
| `check`（别名 `ok`） | 校验 / 放行 | 齐套校验、开工校验、放行卡控、完工确认 | doc、scan |
| `scan` | 扫码 / 采集 | 过站扫码、条码采集、SN 绑定 | check、chip |
| `people` | 人员 / 排班 | 排班派工、人员资质、多能工 | split、clock |
| `gear` | 设备 / 参数 / 建模 | 参数下发、设备配置、四级建模、设备管理 | chip、gauge |
| `chip` | 程序 / 下发 | 程序版本下发、配方下发、DNC | gear、arrow |
| `arrow` | 流转 / 回写 / 派单 | 报工回写 ERP、工序流转、派工 | doc、cycle |
| `down` | 下发 / 上料 | 计划下发、上料、投料指令 | doc、box |
| `clock` | 工时 / ECT | 工时采集、交期测算、节拍、OEE 时间构成 | chart、people |
| `chart` | 排产 / 分析 | 甘特排产、达成率分析、趋势分析 | clock、screen |
| `screen` | 看板 / ESOP | 电子作业指导、车间看板、目视化 | chart、gauge |
| `box` | 物料 / 料箱 | 物料管理、投料、料箱、线边仓 | shelf、truck |
| `shelf` | 三级库存 | 总仓/线边仓/工位仓、库存分层 | box、truck |
| `truck` | 配送 / AGV | 拉动配送、AGV 叫料、送料 | shelf、cycle |
| `warn` | 异常 / 预警 | 异常上报、偏差预警、超期提醒、Andon 触发 | tower、wrench |
| `tower` | Andon 呼叫 | 安灯呼叫、异常响应升级 | warn、people |
| `network` | 集成总线 / 设备联网 | ESB、ERP/PLM/SRM 集成、设备联网采集 | gear、chip |
| `cycle` | 循环拉动 / 双向追溯 | 反向追溯、循环取货、闭环、PDCA | box、shelf |
| `wrench` | 处置恢复 | 返修、异常处置、维修、恢复闭环 | warn、check |
| `route` | 工艺路线 / 过站 | 工艺路线、工序过站、路线切换 | scan、gear |
| `gauge` | SPC / 监控 | 参数监控、SPC/CPK、设备状态仪表 | chart、screen |
| `book` | 知识沉淀 / 复盘 | 知识库、案例沉淀、工程师培养、复盘 | people、screen |
| `factory` | 工厂级 | L3 工厂级 MOM、集团/工厂统一模型 | conveyor、network |
| `conveyor` | 产线级 | L2 产线级执行、线体、工位终端 | factory、gear |
| `bolt` | 拉动信号 | 快速转产、信号触发、换型、共享物料结转 | cycle、truck |
| `calc` | 算法 / 要素数据化 | 排产算法、成本核算、测算引擎、要素建模 | chart、clock |

## 二、按业务模块的成套组合（直接用，不用临场想）

| 业务模块 | 流程链 `flow_icons` | 第二视角卡片图标 |
|---|---|---|
| 生产计划 / 排产 | `doc, check, chart, people, route, arrow` | `calc, chart, clock, screen, bolt` |
| 生产管理 / 工单 | `doc, split, check, people, scan, arrow` | `factory, conveyor, network, gear` |
| 工艺管理 / 防错 | `route, chip, scan, check, screen, cycle` | `book, route, check, screen` |
| 仓库 / 物料拉动 | `box, check, shelf, bolt, truck, cycle` | `shelf, cycle, truck, network` |
| 设备管理 / OEE | `gear, network, gauge, clock, warn, wrench` | `gauge, clock, wrench, chart` |
| 质量管理 / 追溯 | `scan, gauge, warn, wrench, check, cycle` | `gauge, book, network, cycle` |
| 异常管理 / Andon | `warn, tower, people, wrench, check, book` | `tower, warn, wrench, book` |
| 体系总览 / 架构 | `doc, split, check, people, scan, arrow` | `factory, conveyor, network, gear` |

## 三、选型纪律

1. **一页最多出现 8 个不同图标**，重复使用同一图标比硬找新图标更好（语义一致优先）。
2. 流程链图标必须与步骤语义严格对应，禁止整条链都用 `check` 敷衍。
3. 卡片图标不与流程链图标冲突：同一概念在不同区域用同一图标。
4. 图标只做语义提示，**不承担信息**——文字仍须完整表达含义（打印成黑白时不丢信息）。
