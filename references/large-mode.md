# 大额群日订单模式

大额群不是按金额判断，而是固定参与人员、换汇关系较稳定的群。先完整读取
[simple-mode.md](simple-mode.md)：大额群使用其中相同的聊天阅读、订单边界、资金图片、客户、收款方、计价、
核对结果和异常关系规则。本文件只规定大额群额外的选群、单日范围、资金类型与日终统计。

## 自动选群和会计日

默认按统计日期的自然日处理：

```powershell
python scripts/reconcile.py start <原始导出文件或根目录...> --work <全新工作目录> --mode large --date YYYY-MM-DD
```

现场会计日不是自然日时，仍保留 `--date` 作为统计标签，并用精确时间窗覆盖边界，例如：

```powershell
python scripts/reconcile.py start <原始导出文件或根目录...> --work <全新工作目录> --mode large --date 2026-09-03 --from "2026-09-03 05:00" --to "2026-09-04 05:00"
```

- `--mode large` 读取全部群，再排除群名中包含 `小额出` 或 `财务资料群` 的群；名称比较执行统一的 Unicode 和
  空白规范化，不使用白名单，也不逐群点选。
- 必须提供一个 `--date`。未提供时间窗时按 `Asia/Bangkok` 的 `00:00`（含）至次日 `00:00`（不含）筛选；提供
  `--from/--to` 时按该精确区间筛选，起点包含、终点不包含，`--date` 仅作为日终统计标签。
- 所有被选群都要读完并封存。当天没有订单的群不生成空工作表。

## 与小额群相同的完整订单

- 新建大额任务使用 `group-chat-large-daily-decision/2.0`，提交 `orders` 和 `open_orders`，不提交
  `exchanges` 或 `open_exchanges`。
- 一笔完整换汇是一张订单，不是一张图片。付款图、回款图、聊天报价、应回确认和完成消息共同支持同一订单；
  不能按图片数量机械拆单。
- 每张资金图按 [simple-mode.md](simple-mode.md) 的 `media_decisions` 格式记录金额、币种、收款方、状态和发送者
  决定的默认资金方向。每条资金记录必须归入订单或明确的高级关系。
- `amount` 同样只记录收款方实际到账；付款方优惠后实付、立减额或平台补贴金额不进入明细和汇总。实收金额不明时
  按小额模式单图复核并保留待确认。
- 每笔订单必须记录小额群要求的全部内容：订单编号、客户昵称、换汇方向、资金条目、聊天采用的汇率与乘除方向、
  应回依据、实际回款、核对结果和必要异常备注。
- 拆分履约段、合并回款、退款、追回、重复凭证和跨单补抵与小额群使用同一结构；实际出现时读取
  [advanced-relations.md](advanced-relations.md)。
- 大额订单额外必填 `fund_type`，只能是 `wechat`、`alipay`、`bank_card` 或 `usdt`。该字段只用于日终分类统计，
  不能代替资金流水、币种或收款方。
- 跨页未结束订单写入完整 `open_orders`；资金类型已知时一并写入 `fund_type`，未知时在 `unresolved` 中说明。

## 最小批次格式

大额群仍通过 `group-chat-review-batch/1.1` 原子提交：

```json
{
  "contract_version": "group-chat-review-batch/1.1",
  "batch_id": "large-orders-001",
  "base_fingerprint": "sha256:当前指纹",
  "page_commit": {
    "page_start": 0,
    "page_end": 20,
    "page_token": "sha256:review next 返回的页面令牌"
  },
  "open_orders": [],
  "media_decisions": {
    "M0001": {
      "classification": "fund",
      "viewed_original": true,
      "entries": [
        {"amount": "10000", "currency": "CNY", "payee": "wx-account-001"}
      ]
    },
    "M0002": {
      "classification": "fund",
      "viewed_original": true,
      "entries": [
        {"amount": "47200", "currency": "THB", "payee": "123-4-XXX567"}
      ]
    }
  },
  "orders": [
    {
      "id": "L001",
      "entry_ids": ["M0001.1", "M0002.1"],
      "source_messages": ["S00001", "S00002", "S00003"],
      "fund_type": "wechat",
      "customer_nickname": "Alice",
      "direction": "CNY->THB",
      "pricing": {
        "source_messages": ["S00001", "S00003"],
        "terms": {"rate": "4.72", "operator": "multiply"},
        "expected": {"kind": "explicit", "amount": "47200"}
      }
    }
  ]
}
```

`orders` 按 `id` 新增或替换，删除使用 `remove_order_ids`。含 `page_commit` 的批次必须提交完整
`open_orders`；大额 `2.0` 批次出现 `exchanges` 或 `open_exchanges` 会被拒绝。

## 封存、明细与日终统计

```powershell
python scripts/reconcile.py review <work> next --group <群>
python scripts/reconcile.py review <work> apply-batch --group <群> --input <批次.json>
python scripts/reconcile.py review <work> seal --group <群>
python scripts/reconcile.py finish <work> -o <新的大额群日订单.xlsx>
```

封存条件与小额群相同：必须读完群聊、清空 `open_orders`、分类全部可用媒体、查看资金原图、记录可见收款方，
并把资金条目完整归入订单或受控高级关系。聊天真实缺失的客户、方向、应回依据或资金关系可以按小额规则显示为
“待确认”，不能按零或实际回款倒推。

每个有订单的群生成一张分表。明细固定使用小额群相同的 14 列：记录类型、订单编号、客户昵称、换汇方向、
付款合计、汇率、收款方实际到账金额、流水币种、应回金额、内部实际回款合计、核对结果、备注、收款方、聊天消息时间。
订单汇总行下面保留每条客户付款、内部回款、退款、追回及其他小额模式会显示的流水。

明细下方保留独立日终统计区，按 `资金类型 + 换汇方向 + multiply/divide + 精确汇率` 分组，记录笔数、换出合计
和换入合计。同一资金类型使用多个汇率时必须拆行，不计算平均汇率，不把不同币种直接相加。只有计价和资金关系
已经核清的订单或履约段进入日终统计；待确认、失败或取消事实仍保留在资金判定和必要的订单备注中，但不进入合计。

日终统计区必须有独立的深色“资金汇总”标题栏；列标题使用蓝底白字，数据行使用交替浅色底纹和完整边框，笔数、
换出合计、换入合计加粗并使用清晰的数字格式。不要只沿用普通明细的浅色表头。

最终工作簿只包含有订单的实际群分表。不得增加“期间说明”、封面、来源说明、跨群汇总或其他辅助分表；时间范围和
来源说明放在对用户的交付文字中。明细与日终汇总写成静态值，工作簿不得依赖 Excel 专有公式、宏或外部链接，确保
可由 WPS 直接打开和复核。

## 旧工作目录兼容

`group-chat-large-daily-decision/1.0` 的 `exchanges/open_exchanges` 只用于继续既有工作目录及重新生成其旧式工作簿。
新任务不得创建 1.0 判定，也不得把 1.0 与 2.0 判定混在同一次任务中；需要新格式时重新 `start`。
