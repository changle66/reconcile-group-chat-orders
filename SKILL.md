---
name: reconcile-group-chat-orders
description: >-
  从 Telegram、WhatsApp 或 LINE 群聊及备份核对换汇订单：客户发出的资金图默认是付款，内部人员发出的资金图
  默认是回款；小额群和大额群都按完整订单记录客户、资金流水、收款方、方向、汇率、应回依据和异常，
  大额群另按单日汇总资金类型、方向、精确汇率和金额；独立财务资料模式从名称含“财务资料群”的群聊提取
  护照等身份证件及聊天平台账号资料，按人去重合并并保留原图；门店开票群模式把内部人员票据中的 `+/-`
  记为本门店收支，处理引用更正、票据生命周期、跨店调拨和每日余额核对。脚本按实际群生成每群一张分表的 Excel。
  适用于全新处理、继续同一次中断任务或从已封存判定重新生成工作簿；不从旧台账反推聊天事实。
---

# 群聊订单核对

处理小额群前完整读取 [references/simple-mode.md](references/simple-mode.md)；处理大额群前依次完整读取
[references/simple-mode.md](references/simple-mode.md)和[references/large-mode.md](references/large-mode.md)；处理
财务资料群前完整读取 [references/finance-materials.md](references/finance-materials.md)。订单模式出现拆分、多单合并
回款、退款、追回、重复凭证或跨单补抵时，才读取 [references/advanced-relations.md](references/advanced-relations.md)。
处理名称含“门店开票群”的内部货币流水前完整读取 [references/store-ledger.md](references/store-ledger.md)。
涉及 LINE 原始 iOS 备份时读取 [references/line-ios.md](references/line-ios.md)；涉及小米/MIUI LINE Android 应用备份
时读取 [references/line-android-miui.md](references/line-android-miui.md)。任何模式开始查看媒体前完整读取
[references/media-viewing.md](references/media-viewing.md)。

## 当前脚本与 WPS 输出

- 所有命令只从本技能目录运行当前 `scripts/reconcile.py` 及其同目录模块。禁止把脚本复制到任务目录，禁止调用
  `task_local_*`、旧兼容副本或工作区中的同名脚本。继续旧工作目录时也使用当前脚本及其公开升级入口；当前脚本
  无法读取时应停止并修复、测试本技能，不能退回旧脚本绕过。
- 新建任务的 `start` 结果和 `run.json` 必须显示当前 `runtime_contract`、`amount_policy` 与
  `workbook_compatibility`；缺失或不匹配时不得继续录单。
- 最终 `.xlsx` 以 WPS 为主要打开环境：业务数据和汇总写成静态值，不使用单元格公式、宏、外部链接、结构化引用或
  Excel 专有动态数组函数；必要的条件格式仅使用 WPS 支持的基础公式，并由发布检查器验证。

## 群模式

- 默认小额模式保持原流程；`--mode large --date YYYY-MM-DD` 启动大额群日换汇模式。
- `--mode finance` 启动独立财务资料模式，只选择实际群名包含 `财务资料群` 的群。它不建立换汇订单，也不把
  `客户登记（曼谷/芭提雅）`、编号代码或金额写入人员资料；每个实际财务资料群仍各自生成一张分表。
- `--mode store-ledger` 启动门店开票群内部货币流水模式，只选择实际群名包含 `门店开票群` 的群。所有发送者
  只作为审计信息，不按角色推导方向；票据和明确聊天中的 `+` 是本门店收入，`-` 是本门店支出。
- 大额群采用排除式选群：群名包含 `小额出`、`财务资料群` 或 `门店开票群` 时排除，其余群全部自动处理。大额不是按金额判断，
  而是固定参与人员、换汇关系较稳定的群。
- 大额群采用与小额群相同的 `orders`、`open_orders`、资金明细和复杂关系格式；每个大额订单另填 `fund_type`。
  工作表明细使用小额群相同的 14 列，底部再按资金类型、方向、精确汇率和状态统计。已确认与待确认订单都进入
  同一张资金汇总，缺失金额显示“未确认”而不是 0。统计区使用深色分区标题、蓝底白字表头、交替浅色数据行、
  醒目边框和加粗金额，不能与普通明细混在一起难以辨认。

## 订单模式共同口径

- 模型从头到尾阅读每个选中群，判断客户、订单开始与结束、换汇方向、群聊最终采用的汇率、应回依据以及真正需要
  写入订单备注的异常。分页、时间间隔、图片数量、币种和金额接近度都不能代替订单语义。
- 客户发出的普通资金图由脚本记为 `payment`；[config/roster.yaml](config/roster.yaml) 中内部人员发出的普通资金图
  记为 `payout`。成功的内部回款通常表示一单成交并结束，但拆分回款、合并回款、补款、退款、追回或聊天明确继续
  同一单时，仍按上下文处理。
- 普通资金图默认成功。只有图片或聊天明确显示失败、取消、拒绝、作废、无效或风控未完成时才记失败；明确取消的
  订单不能因出现图片而当作成交。
- 资金条目的 `amount` 始终表示收款方实际收到或实际入账的金额。截图同时出现订单原价、优惠、立减、优惠后实付或
  付款方实际支出时，只录收款方实际收到的金额；优惠后实付金额不进入资金流水、订单合计、备注或其他辅助字段。
  收款方实收金额未明确显示或无法区分时，必须单图复核并保留金额待确认，禁止从优惠额、付款方支出或聊天报价推算。
- 模型不需要为普通记录重复填写 `side`、`result`、`amount_state` 或 `payee_state`；脚本按发送者和可见字段生成。
  只有转发付款、代发回款、退款、追回、失败或字段不清楚时才显式填写例外。
- 每个正常订单和每个拆分 `leg` 必须记录群聊采用的汇率、乘除方向和应回依据。正常订单没有明确汇率不能封存为
  已核清；真实缺失时以“待确认”进入表格，汇率列显示“待确认”，不得按实际回款倒推。
- 约定或预期金额、资金凭证显示的转出金额、客户声称的到账金额，以及“齐/OK/完成”等流程结束语是不同证据。
  流程结束语不能单独消除数值冲突；没有明确最终金额或汇率时，对应订单或拆分腿必须保留待确认。

## 看图与收款方

- `review next` 在 `media_queue` 中按时间顺序组织尚未分类的图片。当前主代理按队列批次，用一次编排调用并行打开
  批内图片；不开子代理。订单和门店开票群模式默认每批 9 张并在 4–12 张内自适应，财务资料模式默认 4 张并在 2–6 张内
  自适应；相同哈希只打开一个代表图，缓存命中不重开，缩略图和待复核项单独成批，非图片媒体单列。
- 批内直接查看原图或足以辨认文字的高清图。OCR 候选当前在所有平台默认关闭，需要时允许用户用显式参数手动开启。
  OCR 只处理唯一哈希代表图并按哈希缓存，只能提供候选文字，不能自动写入
  判定、决定订单边界或代替视觉确认；后端不可用或超时就继续直接看原图。
- 每个更新的 `M` 标签在同一 `apply-batch` 中提交一份 `media_observations`。代表图提交完整观察；同批相同图片用
  `reuse_from`，跨页缓存命中用 `reuse_sha256`。观察缓存只保存图上可见事实和复核状态，不保存客户、订单边界、
  汇率、资料归人或归单关系；相同哈希绝不自动表示重复交易。
- 每张资金图都要记录图中明确显示的金额、币种和收款方。任何字段看不清或无法逐图对应时，必须单图放大复核；已经
  清楚的正常图片不重复打开。准确性优先于批量大小和处理速度。
- 泰铢银行转账只记录图片中的收款银行卡号或掩码账号，不记录泰文姓名或银行名；其他币种记录图片中最具体、完整的
  收款姓名、账号或钱包地址。图片确实没有才写 `未显示`，放大后仍看不清才写 `无法辨认`。
- `无法辨认` 的收款方封存前必须定向重看原图；`未显示` 和未知比例偏高只产生告警，不要求重开全部图片。

## 简化工作流

所有命令在本技能目录运行。

### 1. 建立快照

```powershell
python scripts/reconcile.py start <原始导出文件或根目录...> --work <全新且不存在的工作目录> --contains 小额
```

大额群只允许单日任务：

```powershell
python scripts/reconcile.py start <原始导出文件或根目录...> --work <全新且不存在的工作目录> --mode large --date YYYY-MM-DD
```

大额群需要自定义会计日边界时，`--date` 保留为统计日期标签，并同时提供精确时间窗：

```powershell
python scripts/reconcile.py start <原始导出文件或根目录...> --work <全新且不存在的工作目录> --mode large --date 2026-09-03 --from "2026-09-03 05:00" --to "2026-09-04 05:00"
```

财务资料模式默认读取全部日期；可直接输入 Telegram/WhatsApp 导出、含 `Manifest.db` 的未加密 Finder/iTunes
LINE 备份，或小米/MIUI LINE Android 备份：

```powershell
python scripts/reconcile.py start <原始导出文件或备份根目录...> --work <全新且不存在的工作目录> --mode finance
```

门店开票群模式默认读取全部日期；统一入口会自动识别 iOS `Manifest.db` 与 MIUI LINE `.bak`：

```powershell
python scripts/reconcile.py start <原始导出文件或备份根目录...> --work <全新且不存在的工作目录> --mode store-ledger
```

OCR 候选在所有平台默认关闭；手动开启在 `start` 增加 `--ocr-candidates`，手动关闭增加
`--no-ocr-candidates`。解析后的值写入任务清单，同一次任务恢复时保持不变。保持默认关闭时不会导入 OCR 引擎、
初始化模型或启动 OCR 进程。

只处理一个曼谷会计日时，在新任务的 `start` 命令末尾增加 `--date YYYY-MM-DD`，例如
`--date 2026-08-31`。未提供时保持原有的全日期行为；日期不同必须建立新的工作目录。
需要精确时间段时，使用 `--from "YYYY-MM-DD HH:MM" --to "YYYY-MM-DD HH:MM"`；两者都按曼谷时间解释，
开始时间包含、结束时间不包含。小额、财务资料和门店开票群模式不能把时间窗与 `--date` 同时使用；大额模式必须保留 `--date`
作为统计标签，并可用时间窗覆盖默认的 00:00–24:00 边界。

新任务使用全新工作目录；同一次任务中断后继续原工作目录。固定快照、已确认页面、订单模式的 `open_orders`、
财务资料模式的 `open_people`、门店模式的 `open_records`、群判定、语义指纹和批次记录是跨轮次检查点，不得复制到另一项新任务复用。

### 2. 逐群阅读和提交

```powershell
python scripts/reconcile.py review <work> status
python scripts/reconcile.py review <work> next --group <群>
python scripts/reconcile.py review <work> apply-batch --group <群> --input <批次.json>
python scripts/reconcile.py review <work> seal --group <群>
```

- `review next` 默认按实际消息、回复、媒体路径、当前模式的 `open_orders`/`open_people`/`open_records` 和 `carry_messages` 的
  总输出长度，自适应选择预算内
  最大的完整页，并以紧凑 JSON 一次返回；页长不是固定值。它对业务快照和判定只读且绝不推进阅读进度；启用 OCR
  时只可更新可删除的非权威候选缓存。完成该页的语义判断后，
  在同一个 JSON 批次中提交 `page_commit` 和完整的 `open_orders`、`open_people` 或 `open_records`；只有 `apply-batch` 成功才推进
  `reviewed_through`。中断前未提交的页面下次会原样返回，不能形成“没看却已读”。
- 新任务不传 `--limit`，也不在 PowerShell 或工具层把一页拆段显示。显式 `--limit` 只用于继续已经按固定页长启动的
  旧任务或诊断环境异常。只要界面提示截断、JSON 不完整或末尾字段缺失，该页就无效，不能提交。
- 订单模式在一个群内连续阅读和维护订单，跨页未结束订单写入 `open_orders`。财务资料模式按人维护证件和聊天账号，
  跨页尚未完成的人写入 `open_people`；同一本证件或同一账号后来重发时合并回原人。两种模式都要保存起始消息、
  关键消息、媒体、简短事实摘要和仍待确认内容；这些模式没有未结束对象时也必须提交空列表。
- 门店模式按逻辑票据或费用维护 `records`；跨页尚未判定最终版本的对象写入 `open_records`。已经明确为待完成、作废或
  重复的记录应落入 `records` 并写状态原因，而不是永久停留在 `open_records`。
- `apply-batch` 按当前模式执行完整字段和关系校验，成功后原子写入。订单模式另校验角色、资金、收款方和计价；
  财务资料模式另校验证件/账号去重、原图归属和必填字段。本批新增或修改的资金/资料图片计算一次哈希，未改图片
  复用判定中已有的 `evidence_sha256`；媒体观察按内容哈希缓存。`recheck_required` 会生成单图复核队列并阻止
  `seal`，复核后改为 `clear` 或 `rechecked_unreadable`。`check`、`seal` 和 `finish` 仍全量读取复核。日常不再在
  每批后重复运行 `review check`。
- 群尾只定向复核当前模式的摘要、边界、告警和复杂关系；不重新打开已经清楚且无告警的普通图片。随后直接
  `seal`，它会执行最终完整校验。
- `review check` 和 `review audit` 只在批次报错、排查告警或用户要求审计时使用。批量退化指标只提示定向复核，
  不因比例本身永久阻止真实订单封存。
- 不直接修改判定文件，不创建任务专用脚本批量猜订单、金额、收款方、汇率、证件字段或账号归属；通过
  `apply-batch` 保留可恢复的进度和写入保护。

### 3. 发布工作簿

```powershell
python scripts/reconcile.py finish <work> -o <新的群聊订单核对.xlsx>
# 财务资料模式可改用：-o <新的财务资料.xlsx>
# 门店开票群模式可改用：-o <新的门店货币流水.xlsx>
```

`finish` 重新核对快照、已采用的原图哈希和封存指纹，生成临时工作簿并逐格回读，同时执行 WPS 兼容门禁，通过后
才发布。小额模式每个
实际群一张可见分表；大额模式只为当天存在订单的群建分表，并在相同明细格式下方追加日终统计；财务资料模式每个
实际财务资料群一张可见分表，一人一行，在同一行嵌入证件原图和资料页原图。最终工作簿只包含实际群分表，不创建
“期间说明”、封面、来源说明、跨群汇总或其他辅助分表；来源和时间范围写在交付说明中，不占用工作簿标签页。
门店开票群模式同样只创建实际群分表；明细一币种一行，表底按日期和币种核对期初、收支、账面期末与群聊盘点。

## 必须阻止与允许待确认

- 必须阻止：群未读完、仍有 `open_orders`、`open_people` 或 `open_records`、媒体漏分类、资金图或财务资料原图未确认、图片中
  可见的收款方未完整记录、正常订单
  缺少明确汇率、大额订单缺少资金类型、资金重复归单、合并金额不守恒、角色方向与无依据例外冲突、快照或封存
  判定被改写、财务资料图片未归人、相同证件或账号被归给多人、工作簿回读不一致。
- 门店模式另阻止票据/费用/盘点图片未归属、同票号拆成多条逻辑记录、聊天更正没有直接图片引用及完整算式，或门店调拨
  方向与 `+/-` 不一致。待完成、作废、跨店未匹配和余额差额允许封存，但必须原样显示并且不得自动入账或平账。
- 可以封存为待确认：聊天本身无法确认客户、订单方向、应回依据、资金关系，或无法可靠确认资料页属于哪一本证件。
  待确认必须显示具体原因；未知不能按零处理，账号资料也不能因紧邻某张证件图而强行归人。
- 订单备注只写取消、失败、少转、多转、关系不清或其他确实需要读表人知道的问题。正常成功、页面状态、普通费用、
  舍入和处理过程不写备注。

## 修改本技能后的验证

新增或修改业务规则时，测试必须通过当前公开路径：小额判定使用 `group-chat-decision/3.2`，新建大额任务使用
`group-chat-large-daily-decision/2.0`，两者都以 `orders/open_orders` 通过 `review apply-batch` 提交，并使用
`small-group-simple-plan/2.0` 核算；财务资料模式使用 `group-chat-finance-materials-decision/1.0`，以
`people/open_people` 和 `media_decisions` 通过同一 `review apply-batch` 提交。大额 `1.0` 的
`exchanges/open_exchanges` 只保留旧工作目录兼容测试。
门店模式使用 `group-chat-store-ledger-decision/1.0`，以 `records/open_records/balance_snapshots` 和
`media_decisions` 通过同一 `review apply-batch` 提交，并生成 `group-chat-store-ledger/1.0` 静态工作簿。
所有 `simple_ledger` 业务回归测试都使用 plan 2.0；
plan 1.0 只保留一个明确命名的兼容测试。直接修改判定文件的测试仍属于兼容/迁移队列；兼容测试通过不能代替当前
测试通过，也不要继续向该队列增加新业务回归用例。

```powershell
python -X utf8 scripts/run_skill_tests.py current
python -X utf8 scripts/run_skill_tests.py compatibility
python -X utf8 -m compileall -q scripts
python -X utf8 ../.system/skill-creator/scripts/quick_validate.py .
```

`current` 是当前技能发布门槛。`compatibility` 保留尚未迁移的旧入口和旧合同保护；修改共享解析、校验、核算或
工作簿代码时也必须运行。用 `python -X utf8 scripts/run_skill_tests.py inventory` 查看逐项归类；新增、删除或迁移
测试后必须同步更新分区，未分类的数量变化会直接报错。
