---
name: reconcile-group-chat-orders
description: >-
  从 Telegram、WhatsApp 或 LINE 群聊及备份核对换汇订单：客户发出的资金图默认是付款，内部人员发出的资金图
  默认是回款；模型判断订单边界、客户、换汇方向、群聊采用的汇率和少量异常，脚本生成每群一张分表的 Excel。
  适用于全新处理、继续同一次中断任务或从已封存判定重新生成工作簿；不从旧台账反推聊天事实。
---

# 群聊订单核对

开始前完整读取 [references/simple-mode.md](references/simple-mode.md)。只有出现拆分、多单合并回款、退款、追回、
重复凭证或跨单补抵时，才读取 [references/advanced-relations.md](references/advanced-relations.md)。涉及 LINE 原始
iOS 备份时读取 [references/line-ios.md](references/line-ios.md)；涉及小米/MIUI LINE Android 应用备份时读取
[references/line-android-miui.md](references/line-android-miui.md)。

## 核心口径

- 模型从头到尾阅读每个选中群，判断客户、订单开始与结束、换汇方向、群聊最终采用的汇率、应回依据以及真正需要
  写入订单备注的异常。分页、时间间隔、图片数量、币种和金额接近度都不能代替订单语义。
- 客户发出的普通资金图由脚本记为 `payment`；[config/roster.yaml](config/roster.yaml) 中内部人员发出的普通资金图
  记为 `payout`。成功的内部回款通常表示一单成交并结束，但拆分回款、合并回款、补款、退款、追回或聊天明确继续
  同一单时，仍按上下文处理。
- 普通资金图默认成功。只有图片或聊天明确显示失败、取消、拒绝、作废、无效或风控未完成时才记失败；明确取消的
  订单不能因出现图片而当作成交。
- 模型不需要为普通记录重复填写 `side`、`result`、`amount_state` 或 `payee_state`；脚本按发送者和可见字段生成。
  只有转发付款、代发回款、退款、追回、失败或字段不清楚时才显式填写例外。
- 每个正常订单和每个拆分 `leg` 必须记录群聊采用的汇率、乘除方向和应回依据。正常订单没有明确汇率不能封存为
  已核清；真实缺失时以“待确认”进入表格，汇率列显示“待确认”，不得按实际回款倒推。
- 约定或预期金额、资金凭证显示的转出金额、客户声称的到账金额，以及“齐/OK/完成”等流程结束语是不同证据。
  流程结束语不能单独消除数值冲突；没有明确最终金额或汇率时，对应订单或拆分腿必须保留待确认。

## 看图与收款方

- 批量看图是可选的效率优化，不是完成要求。模型或编排脚本可根据信息密度、清晰度和上下文容量选择单图或批量查看；
  脚本只组织候选批次，模型仍可按需拆分。单批不得超过 16 张，不设最低数量。
- 需要精确读取金额、币种、收款方或状态时，优先每批不超过 9 张；10–16 张只用于画面清晰简单或
  `reference/fund` 粗分类。批内直接查看原图或足以辨认文字的高清图。可以用模型视觉或本地 OCR 预读候选文字，
  但 OCR 不能自动写入判定、不能决定订单边界，也不能代替视觉确认。
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

只处理一个曼谷会计日时，在新任务的 `start` 命令末尾增加 `--date YYYY-MM-DD`，例如
`--date 2026-08-31`。未提供时保持原有的全日期行为；日期不同必须建立新的工作目录。
需要精确时间段时，改用 `--from "YYYY-MM-DD HH:MM" --to "YYYY-MM-DD HH:MM"`；两者都按曼谷时间解释，
开始时间包含、结束时间不包含，且不能与 `--date` 同时使用。

新任务使用全新工作目录；同一次任务中断后继续原工作目录。固定快照、已确认页面、`open_orders`、群判定、语义
指纹和批次记录是跨轮次检查点，不得复制到另一项新任务复用。

### 2. 逐群阅读和提交

```powershell
python scripts/reconcile.py review <work> status
python scripts/reconcile.py review <work> next --group <群>
python scripts/reconcile.py review <work> apply-batch --group <群> --input <批次.json>
python scripts/reconcile.py review <work> seal --group <群>
```

- `review next` 默认按实际消息、回复、媒体路径、`open_orders` 和 `carry_messages` 的总输出长度，自适应选择预算内
  最大的完整页，并以紧凑 JSON 一次返回；页长不是固定值。它只读且绝不推进阅读进度。完成该页的语义判断后，
  在同一个 JSON 批次中提交 `page_commit` 和完整的 `open_orders`；只有 `apply-batch` 成功才推进
  `reviewed_through`。中断前未提交的页面下次会原样返回，不能形成“没看却已读”。
- 新任务不传 `--limit`，也不在 PowerShell 或工具层把一页拆段显示。显式 `--limit` 只用于继续已经按固定页长启动的
  旧任务或诊断环境异常。只要界面提示截断、JSON 不完整或末尾字段缺失，该页就无效，不能提交。
- 一个群内连续阅读和维护订单。跨页未结束订单写入 `open_orders`，至少保存起始消息、关键消息、媒体、客户、方向、
  已知汇率、简短事实摘要和仍待确认内容；下一页会同时返回这些状态和对应原消息。没有未结束订单也必须提交空列表。
- `apply-batch` 已执行完整的字段、角色、资金、收款方、计价和关系校验，成功后原子写入。日常不再在每批后重复
  运行 `review check`。
- 群尾只复核订单摘要、相邻订单边界、告警订单和复杂关系；不重新打开已经清楚且无告警的普通资金图。随后直接
  `seal`，它会执行最终完整校验。
- `review check` 和 `review audit` 只在批次报错、排查告警或用户要求审计时使用。批量退化指标只提示定向复核，
  不因比例本身永久阻止真实订单封存。
- 不直接修改判定文件，不创建任务专用脚本批量猜订单、金额、收款方、汇率或备注；通过 `apply-batch` 保留可恢复的
  进度和写入保护。

### 3. 发布工作簿

```powershell
python scripts/reconcile.py finish <work> -o <新的群聊订单核对.xlsx>
```

`finish` 重新核对快照、已采用的资金图哈希和封存指纹，生成临时工作簿并逐格回读，通过后才发布。最终只有一个
工作簿，每个实际群一张可见分表，不创建汇总表。

## 必须阻止与允许待确认

- 必须阻止：群未读完、仍有 `open_orders`、媒体漏分类、资金图未确认、图片中可见的收款方未完整记录、正常订单
  缺少明确汇率、资金重复归单、合并金额不守恒、角色方向与无依据例外冲突、快照或封存判定被改写、工作簿回读
  不一致。
- 可以封存为待确认：聊天本身无法确认客户、订单方向、应回依据或资金关系。待确认必须显示具体原因；未知不能按零
  处理。
- 订单备注只写取消、失败、少转、多转、关系不清或其他确实需要读表人知道的问题。正常成功、页面状态、普通费用、
  舍入和处理过程不写备注。

## 修改本技能后的验证

新增或修改业务规则时，测试必须通过当前公开路径：判定使用 `group-chat-decision/3.2`，语义修改通过
`review apply-batch`，核算使用 `small-group-simple-plan/2.0`。所有 `simple_ledger` 业务回归测试都使用 plan 2.0；
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
