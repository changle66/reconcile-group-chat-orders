# 简单录单模式

公开流程只有 `start → review → finish`。模型阅读完整聊天、查看资金图并判断订单边界、客户、方向、汇率、
应回依据和异常备注；脚本固定来源、生成普通资金方向与状态、校验金额关系、保存检查点并生成工作簿。

## 本次任务与检查点

- 新任务必须使用不存在的 `--work` 目录，不复制旧判定、旧订单或旧工作簿。
- 同一次任务中断后继续原工作目录。`run.json`、固定快照、已确认页面、`open_orders`、群判定、语义指纹和批次记录
  就是跨轮次进度。
- 当前判定合同为 `group-chat-decision/3.2`。3.1 工作目录先运行
  `review <work> upgrade-checkpoints --group <群>`：它保留图片判定和订单，但把不可信的旧阅读游标归零，随后必须
  从第一页重新结合聊天复核。更早合同不能迁移、封存或发布。
- 不编辑 `run.json`、快照、来源指纹、群键或计数。语义事实只通过 `review apply-batch` 提交。

工作目录主要内容：

```text
<work>/
  run.json
  snapshot/normalized.json
  decisions/decision_*.json
```

新任务只处理一个日期时，在 `start` 增加 `--date YYYY-MM-DD`：

```powershell
python scripts/reconcile.py start <原始导出文件或根目录...> --work <全新工作目录> --contains 小额 --date 2026-08-31
```

需要精确时间段时，使用成对的 `--from` 和 `--to`：

```powershell
python scripts/reconcile.py start <原始导出文件或根目录...> --work <全新工作目录> --contains 小额 --from "2026-08-31 05:00" --to "2026-09-01 05:00"
```

日期和时间都按固定业务时区 `Asia/Bangkok` 解释。`--from` 包含起点、`--to` 不包含终点，两者必须同时提供且不能
与 `--date` 同时使用。未提供任何时间参数时行为不变。筛选条件属于固定快照的一部分，改变后必须重新 `start`，
不能修改或复用原工作目录。这些参数只筛选消息时间，不表示按订单完成时间重新归属；订单可能跨出所选区间时，应
不传时间参数，使用全日期快照核对。

## 阅读聊天与查看图片

```powershell
python scripts/reconcile.py review <work> next --group <群>
```

`S00001` 是消息短标签，`M0001` 是媒体短标签。分页不是订单边界。`review next` 是只读操作：重复调用会返回同一页，
不会改变 `reviewed_through`。输出中的 `page_token` 绑定该页的准确内容；处理完成后必须通过 `apply-batch` 提交
`page_commit`，成功后才会进入下一页。最后一页提交成功后 `read_complete=true`，再次查看 `next` 时才有
`done=true`。

默认模式会把本页的消息、回复目标、媒体路径、`open_orders`、`carry_messages`、`page_token` 和
`semantic_fingerprint` 一起计入输出预算，在最多 200 条原始消息中选择预算内最大的完整连续范围，并以单行紧凑
JSON 一次返回。因此每页消息数会随内容长度变化，不需要也不应手工设页长或把一页拆成多段显示。

只有 JSON 完整闭合且末尾同时存在 `page_token` 和 `semantic_fingerprint` 时才能提交。只要工具或界面提示截断、
JSON 无法完整解析或末尾字段缺失，就视为整页未读，不得拿其中可见的令牌推进进度；重新执行同一个 `review next`
应返回同一完整页。显式 `--limit` 保留给已经按固定页长开始的旧任务和环境诊断，不作为新任务的正常流程。

按聊天顺序查看媒体，批量查看是可选的效率优化：

- 模型或编排脚本可根据信息密度、清晰度和上下文容量选择单图或批量查看；脚本只组织候选批次，模型仍可按需拆分。
  单批不得超过 16 张，不设最低数量。
- 需要精确读取金额、币种、收款方或状态时，优先每批不超过 9 张；10–16 张只用于画面清晰简单或
  `reference/fund` 粗分类。
- 批内可以同时打开原图、高清预览或清晰裁图。可用视觉识别或本地 OCR 提供金额、币种、收款方和状态的候选值，
  但候选值必须由模型对图确认后才能提交，不能因图片已经打开就自动确认。
- 批量看清的图片不重复打开；任何字段看不清或无法逐图对应时，必须单图放大复核。准确性优先于批量大小和处理速度。
- 缩略图足以把普通资料图判为 `reference`；资金事实必须来自原图或足以辨认完整文字的高清图。

模型连续维护订单：客户付款通常开始或延续一单，成功的内部回款通常结束一单。拆分、合并回款、补款、退款、
追回、失败或聊天明确继续同一单时，以语义为准。不能按“一张图一单”、时间间隔或金额相近机械划分。页尾仍未结束
的订单写入 `open_orders`；下一次 `review next` 会返回这份状态以及它引用的原消息 `carry_messages`，跨轮次不依赖
模型临时记忆。

## 最小判定格式

普通资金记录只需填写图片事实，方向和正常状态由脚本生成：

```json
{
  "contract_version": "group-chat-decision/3.2",
  "media_decisions": {
    "M0001": {
      "classification": "fund",
      "viewed_original": true,
      "entries": [
        {
          "amount": "100",
          "currency": "CNY",
          "payee": "张三"
        }
      ]
    },
    "M0002": {
      "classification": "fund",
      "viewed_original": true,
      "entries": [
        {
          "amount": "500",
          "currency": "THB",
          "payee": "206-4-XXX781"
        }
      ]
    }
  },
  "orders": [
    {
      "id": "O001",
      "entry_ids": ["M0001.1", "M0002.1"],
      "customer_nickname": "Alice",
      "direction": "CNY->THB",
      "pricing": {
        "source_messages": ["S00001", "S00002"],
        "terms": {"rate": "5", "operator": "multiply"},
        "expected": {"kind": "explicit", "amount": "500"}
      }
    }
  ]
}
```

`evidence_sha256`、来源字段、阅读进度、编辑控制和封存字段由脚本维护。资金条目自动按顺序得到 `M0001.1`、
`M0001.2` 等标识；一张图可以有多条与本单有关的资金记录。

## 普通资金自动规则

- `kind` 默认 `transfer`；现金才填 `kind=cash`。
- `side` 默认按发送者生成：客户为 `payment`，内部人员为 `payout`，角色未知为 `unknown`。
- `result` 默认 `completed`。`status_text` 或 entry `note` 明确出现失败、取消、拒绝、作废、无效或风控未完成时，
  脚本默认 `failed`。处理中、等待确认、待区块确认仍算正常成功过程。
- 同时有金额和币种时 `amount_state` 默认 `clear`；只清楚一项为 `partial`；都不清楚为 `unreadable`。
- 现金 `payee_state` 默认 `cash`；`未显示` 和 `无法辨认` 分别生成 `not_shown`、`unreadable`；其他文字生成
  `visible`。
- 明确例外才填写 `side`、`result`、`amount_state`、`payee_state` 或 `side_exception`。方向例外、退款和追回见
  [advanced-relations.md](advanced-relations.md)。

金额只抄图片明确显示的标准十进制数，币种按图片记录，TRX 不改成 USDT。图片同时显示订单金额和平台或银行承担的
优惠时，记录收款方实际获得的完整订单金额；余额、广告和与本单无关的数字不建资金条目。

## 收款方必须完整

- 每条资金记录必须有收款方。图片中存在的姓名、账号、掩码账号或钱包地址必须完整照录，不能写“截图所示收款方”
  “群内收款方”“泰铢收款账户”等占位文字。
- 泰铢银行转账只记录可见的收款银行卡号或掩码账号，例如 `X-1342`、`XXX-XXX-0116`、`206-4-XXX781`；不记录
  泰文姓名和银行名。账号区域不存在写 `未显示`，放大后仍看不清写 `无法辨认`。
- 其他币种记录图片中最具体、完整的收款姓名、账号或钱包地址；现金由脚本规范为 `现金`。
- `无法辨认` 必须单图重看，并把重看后仍无法辨认的 entry ID 写入 `unknown_payee_reviewed_entry_ids`。
  `未显示` 不重复打开；未知比例高只在 `check/audit/finish` 中告警，不阻止封存。
- 收款方进入 Excel 每条资金明细的“收款方”列，但不参与订单差额计算。

## 订单、汇率与应回依据

普通订单至少填写：

- `id` 和属于本单的 `entry_ids`；一条资金记录只能归入一个订单，合并回款除外。
- `customer_nickname` 与 `direction` 两个键；无法确认时显式留空并保留具体待确认原因。
- `pricing.source_messages`：支持本单最终采用汇率和应回依据的消息标签。
- `pricing.terms.rate` 与 `operator=multiply|divide`：所有正常已核清订单和每个 leg 封存前必填，汇率会写入 Excel。
- `pricing.expected`：群聊最终采用的应回依据。

`order.source_messages` 可以省略。脚本自动合并资金图片消息、计价消息和方向例外消息；模型只在还需补充订单边界、
客户或关系证据时填写额外消息标签。

`pricing.expected` 三种形式：

```json
{"kind": "explicit", "amount": "500"}
```

群里明确最终应回金额时使用；即使公式复算不同，也以明确金额核对，但仍要填写聊天采用的汇率。

```json
{"kind": "calculated_from_terms"}
```

群里明确按完整公式为准时使用，脚本按付款本金、汇率、费用和舍入计算。

```json
{"kind": "unknown", "reason": "not_stated"}
```

真实缺少或冲突时使用；还可用 `conflicting_authority`、`incomplete_formula`。该订单可以封存为待确认，Excel 汇率列
显示“待确认”。不得按实际回款倒推汇率或应回金额。

同一订单或拆分腿出现多个金额时，先区分其业务含义：聊天约定或客户预期是 `quoted_expected`，资金凭证只是
`payout_proof`，客户陈述的到账数是 `claimed_received`，“齐/OK/完成”只表示流程结束。后出现的数字不会仅因时间
更晚而自动覆盖前一个数字；只有明确说明最终金额或明确更正前值的消息，才是 `final_numeric_confirmation`。凭证、
到账陈述或预期金额互相冲突且没有最终数值确认时，`pricing.expected` 使用
`{"kind":"unknown","reason":"conflicting_authority"}`，具体冲突只在订单备注中写一次。

## 受控批次提交

从 `status`、`next` 或报错后的 `check` 取得 `semantic_fingerprint`，创建批次：

```json
{
  "contract_version": "group-chat-review-batch/1.1",
  "batch_id": "tg5459-001",
  "base_fingerprint": "sha256:当前指纹",
  "page_commit": {
    "page_start": 0,
    "page_end": 73,
    "page_token": "sha256:review next 返回的页面令牌"
  },
  "open_orders": [
    {
      "id": "P001",
      "start_message": "S00195",
      "source_messages": ["S00195", "S00199"],
      "media_labels": ["M0040"],
      "customer_nickname": "Alice",
      "direction": "CNY->THB",
      "rate": "5",
      "operator": "multiply",
      "summary": "客户已付款100元，等待内部回款",
      "unresolved": ["等待内部回款"]
    }
  ],
  "media_decisions": {
    "M0003": {"classification": "reference", "note": "收款码资料页"}
  }
}
```

```powershell
python scripts/reconcile.py review <work> apply-batch --group <群> --input <批次.json>
```

- `page_commit.page_start/page_end/page_token` 必须原样来自当前 `review next`。页面必须与已确认进度连续，错误、过期或
  跳页令牌会使整个批次失败且不写文件。
- 含 `page_commit` 的批次必须提交完整 `open_orders`；没有跨页订单时明确写 `[]`。它不是最终订单，不进入 Excel；
  完成、取消或判明无效后，从列表移除，并把真正成交的订单写入 `orders`。
- `open_orders` 至少保存 `id`、`start_message`、关键 `source_messages`、关联 `media_labels`、事实 `summary` 和
  `unresolved`；客户、方向、汇率和乘除方向已经知道时一并保存，未知时不要猜。
- `media_decisions` 按 M 标签新增或替换，`orders` 按订单 ID 新增或替换；未出现的内容保持不变。
- 删除时使用 `remove_media_labels` 或 `remove_order_ids`。
- `balance_links`、`settlement_allocations`、`unknown_payee_reviewed_entry_ids` 出现在批次中时完整替换；省略则保持。
- 脚本在内存合并、补齐普通默认值和证据标签、校验并捕获资金图哈希；失败不写文件，成功后把页面进度、未结束
  订单和判定一起原子替换。
- `apply-batch` 已完成常规校验，不再机械追加一次 `review check`。一个批次覆盖一个读懂的连续片段或若干完整订单，
  不需要按单张图片切批。

禁止直接改判定语义字段，也不创建任务专用脚本批量猜金额、收款方、订单边界、汇率或备注。语义指纹防止中断恢复时
用旧批次覆盖新判断。

## 封存和发布

群尾查看订单摘要和 `apply-batch` 返回的风险指标，只定向复核：

- 相邻订单边界；
- 失败、取消、未知字段或金额差异订单；
- `无法辨认` 的收款方；
- 拆分、合并、退款、追回、重复或跨单补抵；
- 疑似机械“一图一单”的告警订单。

正常且图片文字已清楚的订单不再全量重看。随后运行：

```powershell
python scripts/reconcile.py review <work> seal --group <群>
```

`seal` 要求所有页面均已提交且 `open_orders=[]`，再执行最终完整校验。疑似批量退化只产生风险指标，不因比例本身
永久阻止封存；发现缺少聊天语义、计价或订单关系时，具体缺失内容仍会阻止或进入待确认。

`review check` 用于批次报错排查，`review audit` 用于用户要求的集中审计，均不是每批必跑步骤。

```powershell
python scripts/reconcile.py finish <work> -o <新的群聊订单核对.xlsx>
```

`finish` 要求所有群已封存，重新验证快照、原图哈希、判定指纹、金额关系和汇率字段，生成临时工作簿并逐格回读。
输出已存在时拒绝覆盖。最终每群一张可见分表，包含汇率和每条资金记录的收款方。

复杂关系仅在实际出现时读取 [advanced-relations.md](advanced-relations.md)。
