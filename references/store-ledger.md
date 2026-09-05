# 门店开票群内部货币流水模式

`--mode store-ledger` 只处理实际群名包含 `门店开票群` 的群。参与者均视为内部人员；发送人只用于审计，不能用
roster 或角色推导资金方向。票据或明确聊天中的 `+` 表示该门店收到该币种，`-` 表示该门店支出该币种。

## 输入、平台与日期

```powershell
python scripts/reconcile.py start <原始导出文件或备份根目录...> --work <全新工作目录> --mode store-ledger
```

- 默认读取备份中的全部日期。需要时使用 `--date YYYY-MM-DD`，或成对使用
  `--from "YYYY-MM-DD HH:MM" --to "YYYY-MM-DD HH:MM"`；时间按 `Asia/Bangkok` 解释。
- 同一入口自动发现含 `Manifest.db` 的未加密 Finder/iTunes LINE iOS 备份，以及受支持的 MIUI LINE Android
  `.bak`。平台解析完成后使用同一判定和工作簿合同，不需要手工指定 iOS/Android。
- 日期筛选会保留与区间内消息直接相连的回复链作为 `accounting_context_only` 上下文。区间外上下文不能产生已入账
  流水；已入账记录的 `posting_at` 必须落在本任务时间窗内。
- iOS 加密备份、MIUI 压缩或加密备份、缺失核心数据库会明确停止。无法还原的引用不能按相邻消息猜测。

## 先看引用，再判票据生命周期

`review next` 的每条回复同时显示目标 `S` 标签和目标 `media_labels`。先沿引用关系把内部人员的金额、算式、作废、
改开或取款说明连回具体票据，再结合时间线判断逻辑记录。相近时间、相同金额和相同图片哈希都不能代替引用或明确语义。

- 同票号的未签版本、补签版本、重拍和重复照片属于一个 `record`；用 `media_roles` 标记 `draft`、`final`、
  `duplicate` 或 `supporting`，只选一张非重复图片作为 `representative_media_label`。
- 普通清晰且唯一的票据直接按 `posted` 入账。补签或明确完成的记录以实际完成消息时间作为 `posting_at`；票面日期另存
  `voucher_date`。
- 签收笔迹不要求正好落在印刷的 `Signature` 虚线上。若同票后图新增姓名或签收笔迹，并出现该群反复使用且前后
  一致的完成标记（例如粉色标记），可以判为 `final`；只有空白高亮、订单代号或无法与同票生命周期对应的名字时仍不得入账。
- 明确作废记 `void`；尚未实际完成（例如“12号取现金”）记 `pending`。两者保留票面变动和具体
  `status_reason`，但没有 `posting_at`，实际入账金额留空并排除日汇总。
- 每条记录保存 `status_history`、状态时间、原因和依据消息。跨日 `pending` 次日实际完成后沿用原记录 ID，追加
  `posted` 历史；`voucher_date` 保留票面原日期，`posting_at` 使用实际完成/入账时间。只有 `posted` 金额进入其
  实际发生日余额核对，不能补记回票面日期。
- 作废票据之后发生的独立实际门店调拨是另一条记录，不能为了抵销作废票而把两者合并。
- 同一群中非空 `voucher_number` 只能属于一个逻辑记录；发现同号多版本必须先合并生命周期再封存。

## 图片事实、聊天更正和费用

所有可用图片分类为：

- `voucher`：换汇、现金收支或门店调拨票据；
- `expense`：车票、话费单等费用凭证；
- `balance`：现金盘点或日结余额凭证；
- `reference`：与上述事实无关的图片。

`voucher`、`expense`、`balance` 必须打开原图并在 `facts` 中只记录当前图片可见的票号、票面日期、说明、货币变动、
汇率、泰铢折算或余额。图片事实缓存不得加入聊天中的更正值、记录状态、调拨关系或归单判断。清晰观察至少包含一个
可见事实；模糊、裁切、冲突或字段映射不清时按 [media-viewing.md](media-viewing.md) 进入单图复核。

货币变动的 `basis`：

- `ticket`：采用票据可见值；
- `direct_reply_correction`：只在一条消息直接引用该记录图片，并写出完整的“数值 ×/÷ 数值 = 数值”算式时使用；
  `source_messages` 必须指向这条回复，`note` 说明更正内容。LINE/iOS 键盘产生的 `✖️`、`🟰`、全角等号等同类
  运算符会先归一化；只写结果、间接转述或与图片无引用关系时不得更正；
- `expense_semantics`：车费、话费等语义和凭证足够清晰，即使文字没有写 `-`，仍按本门店支出；
- `chat_explicit`：无票据但聊天明确写出币种、金额和 `+/-` 的实际收支。

有冲突但达不到聊天更正门槛时，整条记录保留 `pending`，不能择一猜测。金额以正数幅度存入 `amount`，方向只由
`sign` 保存；工作簿的实际入账金额再派生为收入正数、支出负数。

## 判定合同

新任务使用 `group-chat-store-ledger-decision/1.0`，仍通过 `group-chat-review-batch/1.1` 原子提交。批次可以更新
`records`、`balance_snapshots`，删除时分别用 `remove_record_ids`、`remove_balance_snapshot_ids`。提交页面时必须带
完整的 `open_records`。

```json
{
  "contract_version": "group-chat-review-batch/1.1",
  "batch_id": "store-page-001",
  "base_fingerprint": "sha256:review-next返回值",
  "page_commit": {
    "page_start": 0,
    "page_end": 20,
    "page_token": "sha256:review-next返回值"
  },
  "open_records": [],
  "media_decisions": {
    "M0001": {
      "classification": "voucher",
      "viewed_original": true,
      "facts": {
        "voucher_number": "0040099",
        "voucher_date": "2026-09-04",
        "movements": [
          {"currency": "USD", "sign": "+", "amount": "100", "rate": "32.65"},
          {"currency": "THB", "sign": "-", "amount": "32.65"}
        ]
      }
    }
  },
  "media_observations": {
    "M0001": {
      "contract_version": "group-chat-media-observation/1.0",
      "classification": "voucher",
      "review_status": "clear",
      "viewed_original": true,
      "recheck_reasons": []
    }
  },
  "records": [
    {
      "id": "R001",
      "record_type": "exchange",
      "posting_status": "posted",
      "posting_at": "2026-09-04T10:01:00+07:00",
      "voucher_number": "0040099",
      "voucher_date": "2026-09-04",
      "source_messages": ["S00001", "S00002"],
      "media_roles": [{"label": "M0001", "role": "final"}],
      "representative_media_label": "M0001",
      "movements": [
        {
          "currency": "USD",
          "sign": "+",
          "amount": "100",
          "rate": "32.65",
          "basis": "ticket",
          "source_messages": ["S00001"]
        },
        {
          "currency": "THB",
          "sign": "-",
          "amount": "3265",
          "basis": "direct_reply_correction",
          "source_messages": ["S00002"],
          "note": "直接引用票据的完整算式更正票面笔误"
        }
      ]
    }
  ],
  "balance_snapshots": []
}
```

`record_type` 只能是 `exchange`、`cash_movement`、`expense`、`internal_transfer`；`posting_status` 只能是
`posted`、`pending`、`void`。一条记录有多少种货币，就在 `movements` 中逐币种记录多少项。

跨页尚未看完生命周期的对象放入：

```json
{
  "id": "R-open-001",
  "source_messages": ["S00020"],
  "media_labels": ["M0008"],
  "summary": "未签票据，等待后续补签或作废说明",
  "unresolved": ["最终状态", "实际完成时间"]
}
```

群读完时 `open_records` 必须为空。已经判明为 pending/void 的对象属于完整 `records`，不是开放对象。

需要次日续接时使用 [roll-forward.md](roll-forward.md)。第二日对前一日票据的直接回复会连回累计快照中的原消息和
原图；更正仍必须直接引用原票并给出完整算式，不能因跨日降低证据门槛。

## 门店调拨

`internal_transfer` 必须增加：

```json
{
  "transfer_id": "T-20260904-001",
  "counterparty_group_name": "芭提雅门店开票群",
  "direction": "send"
}
```

`send` 的全部 movement 必须是 `-`，`receive` 必须是 `+`。两店各自按本群方向入账；发布时脚本按同一
`transfer_id`、相反方向、对方群名及逐币种金额自动判断 `matched/unmatched/mismatch`。匹配流水逐币种跨群净额应为
零；不匹配只显示警告，不补造对方记录，也不自动平账。未实际入账的调拨显示 `not_posted`。

## 余额快照与每日核对

群聊盘点使用：

```json
{
  "id": "B-20260904-close",
  "date": "2026-09-04",
  "kind": "closing",
  "source_messages": ["S00150"],
  "media_labels": [],
  "balances": [
    {"currency": "USD", "amount": "1100"},
    {"currency": "THB", "amount": "6735"}
  ],
  "note": ""
}
```

- `checkpoint` 只保存盘中检查点，不参与正式期末；每群每天最多一个 `closing`，选择当日最后一条完整日结。
- 只记录日结明确报告的币种，缺少币种不能补零。
- 仅当前一自然日有正式 `closing` 时，才把其中已报告的币种滚入下一日期初；首日、断日或上一日缺少正式日结时，
  期初显示“待确认”，不从期末反推。
- `posted` 记录进入收入、支出与净变动；`pending`、`void` 和重复图片不进入汇总。
- 账面期末为期初加净变动。与群聊期末不一致时保留差额并红色警示；缺期初或缺期末也警示。任何情况都不生成调整项。

## 工作簿

每个实际门店开票群一张可见分表，标签使用实际群名；不创建封面、辅助表或跨群汇总表。

- 明细按实际时间排列，一种货币变动一行。列为：入账时间、票面日期、票号/记录号、状态、流水类型、币种、收支、
  符号、票面/说明金额、实际入账金额、汇率、折算泰铢、对方门店、调拨编号、发送人、证据消息、判定依据/异常、
  调拨匹配、代表图。
- 同一逻辑记录只在第一条 movement 行嵌入一张代表图；草稿、重复重拍和其他证据仍保留在判定及来源哈希中。
- 表底按日期和币种列出期初、收入、支出、净变动、账面期末、群聊期末、差额和核对状态。数据全部是静态值，
  无公式、宏或外部链接，并由发布检查器逐格回读。
