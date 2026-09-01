# 少见资金关系

仅在聊天明确出现转发、退款、追回、重复凭证、拆分、多单合并回款或跨单补抵时使用本文件。普通订单不要填写
这些结构，也不要因为金额接近就推断关系。

## 方向例外

普通方向由发送者自动生成。内部人员代客户转发付款图、客户代转内部回款图、退款或追回时，在对应 entry 填写：

```json
"side_exception": {
  "kind": "relayed_customer_payment",
  "source_messages": ["S00123"],
  "detail": "内部人员代客户转发付款凭证，聊天明确说明该图属于客户付款"
}
```

`kind` 只允许：

- `relayed_customer_payment`
- `relayed_internal_payout`
- `explicit_payment_refund`
- `explicit_recovery`

引用必须属于本群，且至少包含一条不是资金图本身的聊天消息。脚本会把这些引用自动并入订单证据。

## 同一交易的重复截图

```json
"same_transactions": [
  {"entry_id": "M0003.1", "same_as": "M0002.1"}
]
```

只有聊天和图片足以确认是同一笔交易时才声明。相同文件哈希、相同金额或相近时间都不自动代表重复。

## 一笔付款拆成多个方向或履约段

至少两个 `legs`，每个 leg 分别填写方向、付款分配、计价和对应回款：

```json
"legs": [
  {
    "leg_id": "thb",
    "display_label": "USDT->THB（转账）",
    "direction": "USDT->THB",
    "allocation_amount": "600",
    "pricing": {
      "source_messages": ["S00100"],
      "terms": {"rate": "32.5", "operator": "multiply"},
      "expected": {"kind": "calculated_from_terms"}
    },
    "payout_entry_ids": ["M0010.1"],
    "recovery_entry_ids": []
  },
  {
    "leg_id": "cny",
    "display_label": "USDT->CNY",
    "direction": "USDT->CNY",
    "allocation_amount": "400",
    "pricing": {
      "source_messages": ["S00105"],
      "terms": {"rate": "7.2", "operator": "multiply"},
      "expected": {"kind": "explicit", "amount": "720"}
    },
    "payout_entry_ids": ["M0011.1"],
    "recovery_entry_ids": []
  }
]
```

所有 leg 的 `allocation_amount` 合计必须等于本单净付款；同一回款记录不能同时属于两个 leg。

`direction` 只表达币种方向。相同方向因汇率、现金/转账等履约方式，或独立讨论的分段结算依据不同而拆成多个 leg
时，每个同方向 leg 都必须填写互不重复的 `display_label`，例如 `USDT->THB（转账）`、`USDT->THB（现金）`。
脚本用该标签生成 Excel 的方向、汇率、核对结果、备注和资金明细；计算仍只使用规范化的 `direction`。

客户把同一币种方向再次明确拆成若干本金段，且各段有各自的凭证、报价或结算争议时，也分别建 leg，例如
`USDT->TRX（首段3U）` 与 `USDT->TRX（后段17U）`。若只是同一口径下的多次回款，不机械拆 leg，多个回款条目仍可
归入同一 leg。无论层次如何，全部 `allocation_amount` 必须与本单净付款严格守恒。

## 数值冲突与加密货币到账

- 约定或预期金额、凭证转出金额、客户声称到账金额和流程结束语分别保留，不相加、不互相替换，也不按“最后出现”
  自动选定权威。
- “齐”“OK”“完成”等只能说明操作流程结束。此前若有金额冲突，除非后续消息明确确认最终金额或明确更正前值，
  仍以 `expected.kind=unknown`、`reason=conflicting_authority` 保留待确认。
- 加密货币凭证记录截图显示的实际转出币种和金额；客户声称的净到账另作为聊天事实。没有明确费用说明时，不推断两者
  差额是网络费、补发、替换或重复，也不按转出或到账结果倒推汇率。
- 待确认备注只写一次具体事实，例如“17 USDT 部分同时出现预期49 TRX、声称到账30 TRX和转出凭证44 TRX，未见
  最终数值确认”；不要再重复脚本已经显示的通用待确认句。

## 一张回款图结清多个订单

使用群级 `settlement_allocations`：

```json
"settlement_allocations": [
  {
    "entry_id": "M0224.1",
    "allocations": [
      {"order_id": "O055", "amount": "9780"},
      {"order_id": "O056", "amount": "20220"}
    ],
    "source_messages": ["S00621"]
  }
]
```

分配额必须为正且严格等于原始凭证金额；目标订单必须是同一已确认客户、同一回款币种，并各自有付款证据。

## 跨单补抵

```json
"balance_links": [
  {
    "source_order_id": "O001",
    "target_order_id": "O002",
    "kind": "shortfall_carryover",
    "amount": "50",
    "currency": "THB",
    "source_messages": ["S00510"],
    "already_in_expected": false
  }
]
```

`kind` 为 `shortfall_carryover` 或 `overpayment_carryover`。只有聊天明确说前单差额补到或抵扣后单时使用；无法
守恒、客户或币种不一致时不要自动应用，两单保留待确认原因。

## 费用与舍入

费用和舍入只按聊天明确采用的规则放入相应 `pricing.terms`。只有客户明确承担的网络费才记录
`network_fee`，并填写 `customer_requested=true`；截图自身展示的矿工费、Gas 或能量消耗不另建资金记录。
