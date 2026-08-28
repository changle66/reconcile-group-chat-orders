---
name: reconcile-group-chat-orders
description: >-
  从 Telegram、WhatsApp 或 LINE 群聊中识别客户付款图片和内部人员回款图片，按订单配对金额、币种与收款方，
  如实记录退款、追回、拆分方向、费用和明确跨单补抵，并生成每群一张分表的 Excel 台账。
---

# 群聊订单核对

本技能从完整群聊中一次性完成图片识别、订单边界和业务语义判定，再由脚本确定性核算并生成 Excel。
开始前完整读取 [references/simple-mode.md](references/simple-mode.md)。

## 职责边界

- 模型逐群从第一条读到最后一条，查看全部证据媒体，并在同一份群判定中完成媒体分类、订单归属、客户、
  方向、采用汇率、应回金额及明确资金关系的判定。每笔订单用 `source_message_ids` 引用支持这些判定的
  同群原消息；引用不受消息距离或分页边界限制。
- 脚本负责原始导出归一化、完整分页、回复解析、指纹与引用校验、媒体覆盖、金额运算、订单编号和 Excel
  输出。脚本不根据关键词、固定消息窗口或金额相近度推断订单语义。
- 模型只抄图片或聊天中明确可见、明确采用的事实，不自行计算金额。收款方记录凭证中最具体的可见收款
  标识；它是证据事实，不是“内部收款方/客户收款方”这类资金角色。只有泰国银行卡转账凭证以收款
  银行卡号作为收款方并保留原图掩码，例如 `206-4-xxx781`；其他付款方式仍记录各自可见的具体收款名称、
  商户或地址。金额求和、
  乘除汇率、费用、舍入、差额和守恒校验全部交给脚本。

## 不变量

- 业务时间统一为 `Asia/Bangkok`；内部人员只由 [config/roster.yaml](config/roster.yaml) 判断，客户标识
  保留平台原值。
- 每个选中群必须完整读取；每张图片或文件凭证必须分类。语音、视频、贴纸和动图不作为资金证据。
- 本技能不包含 OCR 引擎或外部 OCR API。直接查看原图，只抄清楚可见的金额、币种、收款方和状态。
- `unknown != 0`。任一侧无法确认时，相关合计、应回金额和差额留空并显示“待确认”。
- 金额和币种清楚时即可入账；空白、pending 或没有“成功”字样不等于零或失败。只有页面明确失败才按
  失败未计。
- 每张资金凭证都必须明确填写非空 `payee`，后续编译、订单和 Excel 原样传递；未显示写 `未显示`，无法
  辨认写 `无法辨认`，现金写 `现金`。不得根据发送者角色、事件类型或资金方向生成收款方。
- 退款、追回、重复交易、拆分方向、费用、舍入和跨单补抵都是群判定中的结构化数据，不另建关键词规则。
- 最终只有一个工作簿，每个群一张可见分表，不创建汇总总表。

## 工作流

以下命令均在本技能目录运行。用户要求“全新”或“不使用旧数据”时，新建本次工作目录并从原始导出重新
归一化；不要复制旧 decision、events、plan 或 orders。

### 1. 归一化聊天

Telegram 或 WhatsApp：

```powershell
python scripts/normalize_exports.py <导出文件或目录...> -o <work>/normalized.json --timezone Asia/Bangkok --roster config/roster.yaml --force
```

涉及 LINE 时读取 [references/line-ios.md](references/line-ios.md)，使用 `extract_line_ios.py`；多个平台结果用
`merge_normalized.py` 合并。

### 2. 创建每群唯一判定文件

```powershell
python scripts/group_decisions.py prepare <work>/normalized.json --decisions <work>/decisions
```

输出的每个 `decision_*.json` 同时包含该群全部 `media_decisions`、`orders` 和 `balance_links`。只编辑这些
语义字段，不改写群键、消息/媒体标识、数量或指纹。

### 3. 完整读取并判定每个群

```powershell
python scripts/group_decisions.py read <work>/normalized.json --group-key <group_key> --limit 100
```

首次不传 `--cursor`；之后把返回的 `next_cursor` 原样传入，直到 `done=true`。每页按时间顺序返回完整消息、
媒体路径和已解析的直接回复；回复目标即使在另一页也会随当前消息返回。分页只是传输方式，不是订单上下文
边界。读完整个群后填写同一份群判定：

- `media_decisions`：可用证据填写 `order_evidence|reference|uncertain`、资金类型、图片 OCR 和必要的
  `flow_side`；缺失媒体保持空判定。
- 使用脚本批量填写资金凭证时，调用 `group_decisions.fund_evidence_decision(...)`；其 `payee` 没有默认值，
  必须逐图明确传入。不要在任务工作区重新实现会按角色或资金方向推导收款方的 `evidence()` 帮助函数。
- `orders`：用 `event_ids` 归集资金事件，并用 `source_message_ids` 引用报价、公式、客户身份、交接或其他
  支持订单判定的同群消息；方向、汇率、应回金额和其他明确关系也在这里一次填完。
- 所有 `order_evidence` 必须恰好进入一个订单；不能唯一归属的资金图使用 `uncertain`，由台账显示为待归单。

字段契约和高级结构见 `simple-mode.md`。不要先按图片窗口组单，也不要生成订单后再补汇率或应回金额。

### 4. 一次编译事件和完整计划

```powershell
python scripts/group_decisions.py compile <work>/normalized.json --decisions <work>/decisions --events <work>/events --plan <work>/simple_plan.json
```

编译器验证当前源指纹、逐群与逐媒体覆盖、同群消息引用、事件唯一归单和输出目录纯净度，然后同时生成
events 与已经包含全部模型语义的 `simple_plan.json`。该计划是确定性编译产物，不再人工补充或改写。

### 5. 生成订单

```powershell
python scripts/simple_ledger.py <work>/normalized.json --events <work>/events --plan <work>/simple_plan.json -o <work>/orders.json --force
```

### 6. 生成并检查工作簿

```powershell
python scripts/build_workbook.py <work>/orders.json -o <work>/群聊订单核对.xlsx --template assets/模版.xlsx --force
python scripts/check_workbook.py <work>/群聊订单核对.xlsx <work>/orders.json
```

只有 checker 输出 `OK` 的工作簿可以交付。

## 验证

修改本技能后至少运行：

```powershell
python scripts/test_group_decisions.py
python scripts/test_simple_ledger.py
python -X utf8 ../.system/skill-creator/scripts/quick_validate.py .
```

回归测试围绕接口不变量：分页完整且无重复、跨页回复可解析、远距离消息可引用、旧源和跨群引用被拒绝、
每张证据媒体恰好分类、每个订单事件唯一归属、未知不当零，以及所有核算和工作簿检查保持确定性。
