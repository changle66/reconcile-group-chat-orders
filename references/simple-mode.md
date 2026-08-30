# 简单录单模式

本模式只有三个公开阶段：`start` 建立本次快照，`review` 逐群录入，`finish` 核算并发布。模型负责阅读聊天、
划分订单和识别明确的特殊关系，脚本负责来源固定、字段校验、精确运算和工作簿核验。普通资金凭证按发送者
身份确定方向，除明确失败或无效外按成功记录；脚本不根据分页、时间间隔、关键词、金额接近度或图片哈希
划分订单。

## 本次任务与旧数据

- 新任务必须使用一个不存在的 `--work` 目录运行 `start`。不复制旧判定、旧事件、旧订单或旧表格。
- 同一次任务中断后，可继续该工作目录；其中的 `run.json`、固定快照和判定文件是本次进度，不是跨任务缓存。
- `start` 只对原始导出文件和 LINE 数据库做一次来源指纹，不逐张计算全部聊天图片的哈希。
- 只有判为 `fund` 的原图在 `review check/seal` 时计算一次 SHA-256，`finish` 再核对它是否变化。
  `reference` 图片不计算证据哈希。

工作目录主要内容：

```text
<work>/
  run.json
  snapshot/normalized.json
  decisions/decision_*.json
```

`snapshot` 和判定文件都属于当前任务。不要编辑 `run.json`、快照、来源指纹、群键、消息计数或媒体计数。

## 紧凑时间线

```powershell
python scripts/reconcile.py review <work> next --group <群> --limit 200
```

消息按全群时间顺序返回。`S00001` 是本群消息短标签，`M0001` 是本群证据媒体短标签；标签只在当前固定
快照内使用。回复目标即使在其他页也会附带短标签和摘要。系统通知可以折叠，但分页不是订单边界，必须从
第一页读到 `done=true`。

媒体条目中的 `path` 是原图路径。缩略图只能用来判断是否可能是资金凭证；一旦判为 `fund`，必须打开该
路径的原图，看完后立即录入。

## v2.3 群判定

每群只维护一份 `group-chat-decision/2.3` 判定。生成字段保持原样，只编辑
`media_decisions`、`orders` 和确有需要时的 `settlement_allocations`、`balance_links`、
`unknown_payee_reviewed_entry_ids`。示例：

```json
{
  "contract_version": "group-chat-decision/2.3",
  "reviewed_through": 420,
  "read_complete": true,
  "sealed": false,
  "media_decisions": {
    "M0001": {
      "classification": "reference",
      "note": "收款码资料页"
    },
    "M0002": {
      "classification": "fund",
      "viewed_original": true,
      "evidence_sha256": null,
      "entries": [
        {
          "amount": "28890",
          "amount_text": "28,890.00 THB",
          "currency": "THB",
          "payee": "206-4-xxx781",
          "payee_state": "visible",
          "kind": "transfer",
          "side": "payout",
          "result": "completed",
          "status_text": "Transaction successful",
          "amount_state": "clear"
        }
      ]
    }
  },
  "orders": [
    {
      "id": "O001",
      "entry_ids": ["M0002.1"],
      "source_messages": ["S00381", "S00390", "S00402"],
      "customer_id": "user6601768300",
      "customer_nickname": "阿牛",
      "direction": "CNY->THB",
      "rate_state": "not_stated",
      "expected_payout_state": "explicit",
      "expected_payout": "28890"
    }
  ],
  "settlement_allocations": [],
  "balance_links": [],
  "unknown_payee_reviewed_entry_ids": []
}
```

`evidence_sha256` 由 `review check` 或 `review seal` 写入，不手工填写。每张 `fund` 图片中由模型判定与订单
有关的第 1、2 条资金记录，自动得到 `M0002.1`、`M0002.2` 这样的标识。一张图可以录入多条相关记录，但
不是对界面上所有金额做逐项抄录。

## 媒体与资金字段

每个可用证据媒体必须二选一：

- `reference`：账号资料、收款码、报价图、普通照片，或其他不表示资金记录的材料。只保留可选 `note`。
- `fund`：客户或内部人员发出的转账截图、现金票、现金收据、存现单、兑换收据及其他资金凭证。必须有
  `viewed_original=true` 和至少一条 `entries`。现金票由内部人员发出时直接记为现金回款。

资金条目字段：

- `amount`：只抄原图明确金额，使用标准十进制；不得用聊天本金、公式结果或整数替换截图小数。原图同时明确
  显示完整订单金额和平台、银行或福利金抵扣时，按收款方获得的完整订单金额记录，不按付款人优惠后的实付
  金额制造少转差额。
- `amount_text`：可选，原样保留逗号、正负号或单位。
- `currency`：原图币种；TRX 不改写成 USDT。
- `payee`：模型打开原图，识别并原样记录实际收款方最具体的可见姓名、掩码账户或地址；必须是 JSON 字符串，
  账号和钱包地址不得写成数字。
- `payee_state`：必填。明确可见写 `visible`；原图没有收款方字段写 `not_shown` 且 `payee=未显示`；字段存在
  但确实看不清写 `unreadable` 且 `payee=无法辨认`；现金写 `cash`。
- `kind`：`transfer` 或 `cash`；现金的收款方由脚本规范为 `现金`。
- `side`：必填，`payment|payment_refund|payout|recovery|unknown`。普通凭证按发送者身份填写：客户发出记
  `payment`，内部人员发出记 `payout`；聊天明确是退款或追回时分别记 `payment_refund` 或 `recovery`。
- `result`：必填，`completed|failed|pending|not_shown|unknown`。普通凭证默认填 `completed`；只有原图或聊天
  明确显示失败、风控导致未完成、作废、无效、取消或拒绝时才填 `failed`。只有 `completed` 计入合计。
- `amount_state`：`clear|partial|unreadable`。只有 `clear` 必须同时有金额和币种。
- `status_text`：可选，只原样保存页面状态文字，不覆盖 `result`。
- 资金条目通常不填 `note`。正常成功凭证的等待/处理中状态、现金票或收据类型、优惠抵扣、截图金额来源和
  后续“收到/齐”等过程都不写备注；确有订单问题时只在订单级 `note` 说明一次。

普通资金凭证不等待到账、收款或领取确认。“等待确认”“处理中”“待区块确认”等页面文字不属于失败，
仍填 `result=completed`。不能因为聊天里没有后续“收到”或“齐”就改成 `pending` 或漏记。纯风控提示、
无效页面或没有实际资金记录的失败提示仍为 `reference`；有明确金额的失败尝试可记为 `fund/failed`。

客户或内部人员发出的资金凭证都要写入 `entries` 并归入订单。模型主要根据完整聊天划分订单边界、客户、
换汇方向和明确的退款、追回、拆分、补抵或同笔交易关系；界面中与本单无关的余额、优惠、矿工费或广告金额
不另建资金条目。

`payee` 不得使用“截图所示人民币收款方”“截图所示泰国收款账户”“聊天指定USDT收款地址”“群内收款方”
“泰铢收款账户”“USDT收款钱包”等描述性占位内容。脚本校验 `payee_state` 与值一致，并把 Excel 收款方列
强制保存为文本；`未显示` 或 `无法辨认` 不触发订单待确认。

当本群转账条目不少于 10 条且 `not_shown/unreadable` 超过一半时，`review check` 返回
`unknown_payee_review_required=1` 和尚未复核数量。必须逐一重开相应原图，把复核后仍确认为未知的条目标识
完整写入根级 `unknown_payee_reviewed_entry_ids`；`review seal` 不接受批量未知值未经这一步直接发布。

## 订单

每个订单至少填写：

- `id`：本群唯一短编号，例如 `O001`。
- `entry_ids`：属于该订单的资金条目；同一条目不能进入两个订单。被群级 `settlement_allocations` 占用的原始
  合并回款条目不得再出现在任何订单的 `entry_ids` 中。
- `source_messages`：支持订单边界、客户、方向、采用报价或明确关系的消息短标签，可跨页且不限固定半径。
- `customer_id`、`customer_nickname`、`direction`：三个键都必须存在，由模型填写；聊天无法确认时显式写
  `null` 或空字符串。多方向订单的订单级 `direction` 可为空，各 `leg.direction` 必须明确。
- 每个计价范围都必须填写 `rate_state` 和 `expected_payout_state`。普通订单以订单为计价范围；多腿订单以
  每个 `leg` 为计价范围，订单级计价字段与状态必须留空。

`note` 是模型根据完整聊天填写的订单备注。只有订单确有需要向读表人说明的问题时才填写；没有问题就省略。
优惠抵扣、普通费用、舍入方式、现金票类型、凭证页面状态和正常订单衔接过程都不属于问题。备注只陈述聊天
和原图支持的实际异常，不代替资金记录，也不由脚本根据订单结构、金额或币种自动生成。

`rate`、`rate_operator`、`expected_payout` 只在聊天中有依据时填写。模型根据完整对话确定订单边界、客户和
方向；脚本不从发送者推客户，也不从资金币种拼接方向，更不反推汇率或猜测应回。

汇率与应回状态：

- `rate_state=adopted`：聊天明确采用了汇率，必须填写 `rate`；如果要由汇率计算应回，还必须填写
  `rate_operator=multiply|divide`。
- `rate_state=not_stated`：聊天没有明确采用汇率；不得填写 `rate` 或 `rate_operator`。
- `rate_state=uncertain`：无法确认采用的汇率；不得填写 `rate` 或 `rate_operator`。
- `expected_payout_state=explicit`：聊天明确给出最终应回，必须填写 `expected_payout`；它是权威值。
- `expected_payout_state=calculated_from_rate`：聊天明确要求按采用汇率计算应回；不得手填 `expected_payout`，
  必须同时使用 `rate_state=adopted` 并填写 `rate` 和 `rate_operator`，由脚本按现有核算公式计算。
- `expected_payout_state=not_stated`：聊天没有明确应回；不得填写 `expected_payout`，且不能与
  `rate_state=adopted` 并用。
- `expected_payout_state=uncertain`：应回无法确认；不得填写 `expected_payout`，且不能与
  `rate_state=adopted` 并用。

必填的是状态，不是数字。聊天没有依据时必须如实选择 `not_stated` 或 `uncertain`，不能用零、倒算或经验值
补空。`review check` 会报告 `pricing_scopes` 和 `missing_pricing_states`；缺状态时可以继续审阅，但
`review seal` 和 `finish` 都会拦截。明确应回、汇率和运算符同时存在时，脚本仍按原规则复算；结果冲突则
保留明确应回，并把订单标为待确认。

所有资金条目在封存前都必须由模型明确归入一个订单。无法唯一归单时，由模型建立一个显式待确认订单，
填写相应 `entry_ids`，并把无法判断的 `side`、客户或方向写为未知；脚本不会自动造单。

## 少量高级关系

以下字段只在聊天明确出现时使用；普通订单不要填写空结构或人为制造关系。

同一交易的不同截图：

```json
"same_transactions": [
  {"entry_id": "M0003.1", "same_as": "M0002.1"}
]
```

`same_as` 指向实际计数的主记录；同一交易有摘要页和详情页时，优先把信息完整的详情页作为主记录。工作簿
只显示主记录，重复记录不显示也不生成备注。是否为同一交易完全由模型声明：图片哈希相同只说明文件字节
相同，不会自动去重，也不会自动触发跨群待确认。一张图片内与订单有关的多条资金记录不是重复。

一笔付款明确拆成多个换汇方向时，才使用至少两个 `legs`：

```json
"legs": [
  {
    "leg_id": "thb",
    "direction": "USDT->THB",
    "allocation_amount": "600",
    "rate_state": "adopted",
    "rate": "32.5",
    "rate_operator": "multiply",
    "expected_payout_state": "calculated_from_rate",
    "payout_entry_ids": ["M0010.1"]
  },
  {
    "leg_id": "cny",
    "direction": "USDT->CNY",
    "allocation_amount": "400",
    "rate_state": "not_stated",
    "expected_payout_state": "explicit",
    "expected_payout": "2000",
    "payout_entry_ids": ["M0011.1"],
    "recovery_entry_ids": []
  }
]
```

费用和舍入沿用订单或 leg 内的 `fees`、`rounding`。只有聊天明确要求客户承担并从本金或应回扣除的网络费
才记录 `network_fee`，且必须有 `customer_requested=true`；截图自身显示的矿工费、Gas 或能量消耗不录为
额外费用。

一张实际回款或追回凭证明确同时属于至少两个普通订单时，使用群级 `settlement_allocations`：

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

源条目必须是金额、币种明确且已完成的 `payout` 或 `recovery`，只能由这一关系占用；至少分给两个不同订单，
各分配额必须为正且合计严格等于原始凭证金额。目标订单必须是同一客户、同一回款币种的普通订单，并各自有
付款证据。脚本为每个目标订单生成一条“内部回款分摊”明细，显示本单计入额，同时在备注保留原始合并回款
总额；原始物理凭证只计数一次。它与 `legs` 不同：`legs` 是一笔付款拆成多个方向，
`settlement_allocations` 是一笔实际回款结清多个订单。

明确把前单差额补到或抵扣后单时，才使用群级 `balance_links`：

```json
{
  "source_order_id": "O001",
  "target_order_id": "O002",
  "kind": "shortfall_carryover",
  "amount": "50",
  "currency": "THB",
  "source_messages": ["S00510"],
  "already_in_expected": false
}
```

`legs` 是核算结构，不是工作簿行。客户付款、付款退款、内部回款和回款追回仍只显示为相应资金明细，不能
从 `legs` 再生成或改名为“换汇明细”。正确闭合的拆分不要求备注；只有确有问题时才填写订单 `note`。

重复、超额、客户不一致、币种不一致或无法与前单实际差额守恒的补抵关系不自动应用，两单都标为待确认。

## 校验结果

以下属于录入或技术完整性错误，会阻止封存或发布：

- 群未读完、可用媒体漏分类、资金图未打开原图；
- 描述性收款方、字段格式错误、缺少显式 `side`/客户/方向、短标签不存在、资金条目未归单或重复归单；
- 收款方状态和值冲突、批量未知收款方未逐图复核，或合并回款分摊不守恒、跨客户、币种不一致；
- 缺少汇率状态或应回状态，状态值不受支持，或状态与汇率、运算符、明确应回互相冲突；
- 固定快照、已采用资金原图或封存后的语义判定发生变化；
- 输出文件已存在或生成的工作簿回读不一致。

以下属于模型已显式记录的业务未知，相关订单仍进入工作簿并显示“待确认”：

- `side=unknown`、客户为空或方向为空；
- 金额不完整或 `result` 未确定；
- 明确应回与汇率复算冲突；
- 补抵、拆分或其他关系无法由聊天唯一确认。

原始导出已缺失的媒体只进入缺失统计，不被脚本伪造为资金记录或待确认订单。

明确失败的资金明细及 `same_transactions` 声明的重复明细只保留在内部证据中，不生成工作簿明细行，也不因
其本身生成备注。

原则是：技术输入有错就停止，业务事实不完整就如实保留；两者都不能用零或猜测掩盖。

## 发布

```powershell
python scripts/reconcile.py review <work> check --group <群>
python scripts/reconcile.py review <work> seal --group <群>
python scripts/reconcile.py finish <work> -o <新的群聊订单核对.xlsx>
```

`finish` 要求所有群已封存，重新核对快照与采用的资金原图，在临时目录编译订单、生成 Excel 并回读验证，
通过后才原子发布。失败和已声明重复的资金明细不写入工作簿，普通无问题订单不写备注。最终交付只有指定的
`.xlsx`；内部 events、plan 和 orders 不留在工作目录。

旧的 `group-chat-decision/2.2` 仅作已有任务兼容；新任务必须从 `start` 生成 2.3 判定后审阅和封存。
