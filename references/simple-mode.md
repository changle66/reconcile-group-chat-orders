# 简单录单模式

目标是从完整群聊中如实记录资金证据及其订单关系。模型在一份群判定中完成全部语义工作，脚本校验来源并
确定性核算。关系不唯一时保留未知，不补零、不猜测。

## 完整群读取接口

`group_decisions.py read` 对应以下接口：

```text
read_group(group_key, cursor, limit)
  -> messages, next_cursor, done, group_fingerprint
```

- 消息严格按归一化后的时间顺序分页；`next_cursor` 是绑定当前群内容的游标，原样传回即可。
- 从无游标的第一页开始，持续读取到 `done=true`，即可恰好覆盖全群消息一次。
- 每条消息保留 `message_id/source_sequence/timestamp/sender/text/reply/media` 等归一化字段。
- `reply_message` 从全群解析，所以回复目标在其他页时仍可直接看到。
- 分页只解决传输大小，不定义订单边界。模型可引用任意距离的同群消息。
- 群内容变化后，旧游标和旧 `group_fingerprint` 都会失效。

返回页示意：

```json
{
  "contract_version": "group-chat-page/1.0",
  "group_key": "telegram:group",
  "group_fingerprint": "sha256:...",
  "message_count": 1200,
  "page_start": 100,
  "page_end": 200,
  "messages": [],
  "next_cursor": "v1.200....",
  "done": false
}
```

## 群判定接口

`group_decisions.py prepare` 为每个群生成一份 `group-chat-decision/1.0` JSON。它是该群唯一的模型语义
输出，媒体分类和订单字段在这里一起填写：

```json
{
  "contract_version": "group-chat-decision/1.0",
  "normalized_source_fingerprint": "sha256:...",
  "group_fingerprint": "sha256:...",
  "group_key": "telegram:group",
  "media_decisions": [
    {
      "media_id": "telegram:group:101#media:0",
      "event_id": "telegram:group:101#media:0#event",
      "message_id": "telegram:group:101",
      "availability": "available",
      "decision": {
        "disposition": "order_evidence",
        "event_type": "payment_screenshot",
        "flow_side": null,
        "ocr": {
          "amount": "6048",
          "currency": "CNY",
          "payee": "Goddess Space",
          "status_text": "成功",
          "status_class": "completed",
          "status_class_confidence": "high",
          "amount_completeness": "complete",
          "confidence": "high"
        },
        "note": null
      }
    }
  ],
  "orders": [
    {
      "event_ids": [
        "telegram:group:101#media:0#event",
        "telegram:group:118#media:0#event"
      ],
      "source_message_ids": [
        "telegram:group:80",
        "telegram:group:87",
        "telegram:group:101",
        "telegram:group:118"
      ],
      "customer_id": "user6601768300",
      "customer_nickname": "阿牛",
      "direction": "CNY->THB",
      "rate": "4.96",
      "expected_payout": "30000"
    }
  ],
  "balance_links": []
}
```

群键、来源指纹、消息/媒体/事件标识、可用状态和计数由脚本生成，不编辑。编译器要求每个选中群恰好一份
判定、每张证据媒体恰好一条判定、每个 `order_evidence` 事件恰好进入一个订单。

### 媒体判定

先判断 `disposition`，再抄资金内容：

- `order_evidence`：属于当前群某笔订单；填写资金 `event_type/ocr`，需要覆盖默认资金侧时填写
  `flow_side=payment|payment_refund|payout|recovery`。
- `reference`：账号页、收款码、报价图、普通照片或明确属于其他群/平台的展示材料；不得填写资金
  `event_type/flow_side/ocr`。
- `uncertain`：能读出资金内容但无法唯一确定订单；填写资金 `event_type/ocr`，不填 `flow_side`。编译后
  作为可见的未归单资金图进入台账。
- 媒体缺失时整个 `decision` 保持空值，由编译器生成 `missing_evidence_media`。

资金 `event_type` 支持 `payment_screenshot|payout_screenshot|cash_payment|cash_payout|payout_recovery`。
发送者角色只提供默认资金侧；明确的聊天语义用 `flow_side` 覆盖，而不是为某句关键词建立规则。

### 图片字段

- `ocr.amount/currency/payee/status_*` 只来自图片可见内容，不用聊天公式回填 OCR。
- `payee` 必填并保留原图掩码；未显示写 `未显示`，看不清写 `无法辨认`，现金固定写 `现金`。
- 现金的 `amount` 使用无千分位的标准十进制数；`amount_text` 可保留票面的正负号、逗号或手写单位。
- 图片明确失败时使用 `status_class=failed`；没有状态时使用 `blank`。`pending|blank|unknown` 不自动排除
  清楚且完整的金额，只有失败或证据不完整的流水不进入合计。
- 明确显示 TRX 就记录 TRX，不换写成 USDT。

### 订单字段

每个订单必填：

- `event_ids`：该订单的资金事件；同一事件不能出现在两个订单。
- `source_message_ids`：支持订单边界、客户、方向、采用报价、公式或明确资金关系的同群消息。它是证据
  引用，不是固定半径；可以跨页、跨越无关对话并包含资金图片消息。

通用可选字段：

- `case_id`：稳定业务标识；省略则按群、事件与来源消息生成。
- `start_message_id`：订单日期锚点；省略则使用引用消息和资金消息中最早的一条。
- `customer_id/customer_nickname`：客户无法由付款发送者唯一推出时显式填写。
- `direction`：例如 `CNY->THB`、`USDT->CNY`。
- `rate`：聊天中明确采用的纯数值汇率。
- `rate_operator`：`multiply` 或 `divide`；省略为乘法。
- `expected_payout`：聊天明确给出的最终应回金额；提供后视为已经包含声明的费用和舍入，不再二次调整。
- `same_transactions/legs/fees/rounding`：明确关系的结构化数据，格式见下文。

不要把未出现的语义字段补成零，也不要在编译生成 `simple_plan.json` 后再次补写这些字段。

## 编译产物

`group_decisions.py compile` 根据群判定同时生成 events 与完整 `small-group-simple-plan/1.1`。events 只保留
图片资金事实及来源，plan 保存订单关系、引用和业务字段；两者都带当前来源指纹。编译器只验证通用结构：

- 决策与当前 normalized/group 指纹一致；
- 所有群、消息、媒体和事件引用存在且属于同一群；
- 所有证据媒体有且仅有一个分类；
- 所有订单事件有且仅属于一个订单；
- 输出 events 通过币种、收款方、媒体归属与完整覆盖校验。

编译器不从关键词、消息距离或金额近似推断订单、汇率或应回金额。

## 资金净额与反向流水

默认情况下，客户候选发送的资金图为付款，内部人员发送的资金图为回款。明确相反时在对应
`media_decision.decision.flow_side` 中写真实方向。脚本计算：

```text
客户净付款 = 客户付款 - 付款退款
内部净回款 = 内部回款 - 回款追回
```

内部代发客户付款图时，同时在订单填写真实客户；身份仍不唯一则客户留空并显示“待确认”。

## 同一交易的不同截图

完全相同的文件 blob 自动去重。不同文件若明确属于同一交易，在订单中声明：

```json
"same_transactions": [
  {"event_id": "payment-detail", "same_as": "payment-overview"}
]
```

两张图都显示，但只计一次；金额或币种冲突时校验失败。

## 一笔付款的多个换汇方向

聊天明确把一笔净付款拆给多个方向时使用 `legs`：

```json
"legs": [
  {
    "leg_id": "thb",
    "direction": "USDT->THB",
    "allocation_amount": "600",
    "rate": "32.5",
    "payout_event_ids": ["payout-thb"]
  },
  {
    "leg_id": "cny",
    "direction": "USDT->CNY",
    "allocation_amount": "400",
    "rate": "5",
    "payout_event_ids": ["payout-cny"]
  }
]
```

`allocation_amount` 来自聊天中的明确分配，不是图片 OCR。所有分配必须使用同一付款币种并与客户净付款
守恒；每个回款/追回事件只能属于一个 leg。各 leg 可独立使用 `rate/rate_operator/expected_payout/fees/rounding`。

## 费用与舍入

费用必须有种类、金额、币种和处理方式：

```json
"fees": [
  {"kind": "delivery_fee", "amount": "2", "currency": "USDT", "treatment": "added_to_payment"},
  {"kind": "service_fee", "amount": "4", "currency": "THB", "treatment": "deducted_from_payout"}
],
"rounding": {"unit": "5", "mode": "down", "currency": "THB"}
```

`kind` 支持 `delivery_fee|service_fee|network_fee`；`treatment` 支持
`added_to_payment|added_to_payout|deducted_from_payout|included_in_quote|separate`；舍入模式支持
`half_up|down|up`。平台截图自身显示的矿工费、网络费、Gas 或能量消耗不创建费用流水。只有聊天明确要求
代转手续费资产并从本金或应回扣除时才使用 `network_fee`，同时设置 `customer_requested=true`。

## 明确跨单补抵

前单差额明确约定在后单补回或抵扣时，在群级 `balance_links` 记录：

```json
{
  "source_case_id": "first",
  "target_case_id": "second",
  "kind": "shortfall_carryover",
  "amount": "50",
  "currency": "THB",
  "source_message_ids": ["telegram:group:500"]
}
```

`kind` 支持 `shortfall_carryover|overpayment_carryover`；若后单的明确应回已含调整，设置
`already_in_expected=true`。脚本校验客户、币种、时间和前单实际差额；关系不唯一时两单显示“待确认”。

## 未归单、未知与核算结果

- `uncertain` 资金图自动生成可见“未归单资金图片”记录，不进入任何订单合计。
- 失败、看不清、重复和缺失证据都保留可见明细及原因，但不误计。
- 任一核算侧不完整时，该侧合计、应回金额与差额保持空值；未知永远不作为零。
- 两侧完整但没有明确汇率或应回金额时，可以完成资金录入，但不计算多转/少转。
- 只有两侧完整且应回已知时才计算差额；币种容差、费用、舍入、分配守恒和余额关系均由脚本处理。

## Excel 输出

Excel 固定为 15 列：`记录类型、订单编号、客户昵称、客户标识、换汇方向、付款合计、汇率、流水金额、
流水币种、应回金额、内部实际回款合计、核对结果、备注、收款方、聊天消息时间`。订单汇总、换汇明细、
费用和舍入行的收款方可空；所有资金流水行必须非空。每张分表默认冻结首行，表头和数据单元格均水平、
垂直居中；“核对结果”以“少转”开头时使用绿色字体，以“多转”开头时使用红色字体。

群判定是可追溯的唯一模型语义输入；events、plan、orders 和工作簿都是由脚本生成并校验的下游产物。
