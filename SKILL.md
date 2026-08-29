---
name: reconcile-group-chat-orders
description: >-
  从 Telegram、WhatsApp 或 LINE 群聊逐群查看原始资金凭证，按订单核对付款、回款、收款方、退款、追回、
  拆分与明确跨单补抵，并生成每群一张分表的 Excel 台账。适用于全新处理群聊导出、继续同一次核对任务或
  重新生成已封存任务的工作簿；不用于从旧台账反推聊天事实。
---

# 群聊订单核对

本技能按“建立本次快照 → 逐群审阅 → 一次发布”工作。开始核对前完整读取
[references/simple-mode.md](references/simple-mode.md)；涉及 LINE 原始 iOS 备份时再读取
[references/line-ios.md](references/line-ios.md)。

## 职责

- 模型逐群从第一条读到最后一条，查看全部证据媒体，并在同一份群判定中一次完成媒体分类、订单归属、客户、
  换汇方向、采用汇率、应回金额及明确资金关系的判定。每个普通订单、每个拆分 `leg` 都必须同时填写
  `rate_state` 与 `expected_payout_state`；不得先生成订单再补汇率或应回。必填的是状态，不是数值：聊天没有
  明确依据时写 `not_stated` 或 `uncertain`，禁止按实际回款倒推或猜测。
- 模型用缩略图筛选媒体，但资金事实只能在打开原始图片后抄录。图片金额不能被聊天中的本金、公式或整数
  替换。
- 脚本负责原始导出解析、群筛选、阅读进度、短标签映射、证据哈希、字段校验、金额运算、订单编号、Excel
  生成和回读核验。脚本只执行模型显式声明的同笔交易、拆分和补抵关系。
- 模型不计算金额，只抄原图资金事实并划分订单。普通资金方向按发送者身份填写，除明确失败或无效外一律按
  成功记录；只有订单边界和聊天明确出现的特殊资金关系需要模型判断。
- 模型只在订单确有需要向读表人说明的问题时填写订单备注；没有问题就不填。正常成功凭证的页面状态、
  现金票类型、优惠抵扣、普通费用、舍入方式和衔接过程都不写入备注；脚本也不因订单结构本身自动编写
  业务说明。

## 不变量

- 业务时间为 `Asia/Bangkok`；内部人员只由 [config/roster.yaml](config/roster.yaml) 标注。普通资金凭证由客户
  发出就记 `payment`，由内部人员发出就记 `payout`；聊天明确是退款或追回时，分别改记
  `payment_refund` 或 `recovery`。
- 每条 `fund` 记录仍须显式填写 `side=payment|payment_refund|payout|recovery|unknown`，但普通付款与回款直接
  使用上述发送者规则，不再逐笔推断。
- 每个选中群必须读到末尾；每个可用图片或文件凭证必须判为 `reference` 或 `fund`。原始导出已经缺失的
  媒体只计入缺失统计，不由脚本虚构资金记录或订单。语音、视频、贴纸和动图不进入资金证据队列。
- 不使用 OCR 引擎或外部 OCR API。直接查看原图，只抄明确可见的金额、币种、收款方和页面状态。
- 每条资金记录必须有收款方。模型打开原图，识别并原样记录实际收款方最具体的可见姓名、掩码账户或地址；
  未显示写 `未显示`，看不清写 `无法辨认`，现金由脚本写 `现金`。脚本只校验并原样传递，不识别、不推断、
  不补值；收款方不参与订单状态判断。
- `截图所示人民币收款方`、`截图所示泰国收款账户`、`聊天指定USDT收款地址`、
  `固定金额二维码收款方` 等描述性占位内容非法，不能封存群判定。
- 客户或内部人员发出的转账截图、现金票、现金收据、存现单、兑换收据等资金凭证都必须录入；一张图片可以
  录入多条与订单有关的资金记录，界面中与本单无关的余额、优惠或广告金额不另建条目。
- 原图同时显示订单金额和由平台、银行或福利金承担的优惠抵扣时，资金金额按收款方实际获得的完整订单金额
  记录，不把付款人优惠后的实付金额误判成少转；优惠抵扣本身不生成备注。
- 每条已录入资金记录必须由模型明确归入一个订单；脚本不自动创建“未归单订单”。无法唯一判断时，由模型
  建立显式待确认订单并填写未知字段。
- SHA-256 只用于确认采用的原图没有变化，不代表“同一笔交易”。只有模型填写的 `same_transactions` 才会
  去重计数；相同哈希不会自动去重或触发跨群待确认。
- 普通资金凭证默认 `result=completed`。只有原图或聊天明确表示失败、风控导致未完成、作废、无效、取消或
  拒绝时才记 `failed`；“等待确认”“处理中”“待区块确认”不算失败，仍按成功记录。无需等待后续出现
  “收到”“齐”或现金实际领取确认。`status_text` 只原样保存页面文字，不覆盖这个简化口径。
- `unknown != 0`。证据不完整只影响相关合计和差额，不能当作零。
- 每个实际计价范围必须显式填写汇率状态和应回状态。普通订单以订单为计价范围；多方向订单以每个 `leg` 为
  计价范围，订单级不重复填写。`calculated_from_rate` 只声明由脚本计算，模型不得手算应回金额。
- 明确失败的资金尝试和模型已声明为同一交易的重复记录保留在内部证据中，但不生成 Excel 明细行，也不因
  其本身生成备注。`legs` 只用于核算拆分，不得复制或改名为“换汇明细”行。
- 最终只有一个工作簿，每个实际群一张可见分表，不创建汇总表。

## 工作流

所有命令在本技能目录运行。旧的 normalize、group_decisions、simple_ledger、build 和 check 脚本仍是内部兼容
实现；新任务不要直接串联它们，也不要生成任务专用 `fill_decisions.py`。

### 1. Start：建立全新任务

```powershell
python scripts/reconcile.py start <原始导出文件或根目录...> --work <全新且不存在的工作目录> --contains 小额
```

`start` 递归发现原始 Telegram `result.json`、WhatsApp 原始聊天文本和 LINE 备份，排除历史核对目录，按
解析后的实际群名筛选。它创建固定聊天快照、平台前缀群标签和每群一份 v2.2 判定文件。禁止覆盖已有工作
目录；用户要求“全新”或“不使用旧数据”时必须运行新的 `start`，不能复制旧判定。

当前任务中断后可以继续同一工作目录；这是本次任务进度，不是跨任务缓存。

### 2. Review：逐群完成

查看进度或读取下一页：

```powershell
python scripts/reconcile.py review <work> status
python scripts/reconcile.py review <work> next [--group <群键或平台群标签>] [--limit 200]
```

`next` 自动记录阅读位置，以 `S00001` 消息标签和 `M0001` 媒体标签返回紧凑时间线；平台系统通知折叠但保留
在快照中。逐群工作，不混合相同显示名的不同平台群。

审阅时：

1. 连续阅读时间线并维护本群订单草稿；分页不是订单边界。
2. 缩略图只做 `reference/fund` 筛选。凡是 `fund`，使用图片查看工具打开 `path` 指向的原图，并结合完整
   聊天判断其中哪些资金事实属于订单。
3. 看完原图立即把事实写入该群 `decision_*.json`；建立订单或拆分 `leg` 时同时完成汇率状态和应回状态，
   不要等所有群看完后集中回填。
4. 每批录入后运行 `review check`，让脚本规范字段、捕获新资金证据哈希，并检查返回的
   `missing_pricing_states`；封存前必须降为零。
5. 读到群尾、媒体分类完整且订单关系已复核后运行 `review seal`。

```powershell
python scripts/reconcile.py review <work> check --group <群键或平台群标签>
python scripts/reconcile.py review <work> seal --group <群键或平台群标签>
```

使用 `apply_patch` 直接更新现有 v2 判定文件的语义字段。不要编写临时 Python 帮助函数批量生成收款方、
金额或方向；`check` 和 `seal` 是唯一允许补充生成字段的接口。字段契约与示例见 `simple-mode.md`。

### 3. Finish：核算并发布

```powershell
python scripts/reconcile.py finish <work> -o <新的群聊订单核对.xlsx>
```

`finish` 要求所有群已封存，并重新检查聊天快照和采用的资金原图。它在内存中编译事件、计划和订单，生成
临时工作簿，再逐格回读核对；只有核对通过才发布最终文件。工作目录中不落盘 `events/`、
`simple_plan.json` 或 `orders.json`，输出文件已存在时拒绝覆盖。

业务不明确的订单以“待确认”进入工作簿，不阻止其他可靠订单交付；快照损坏、群未读完、媒体漏分类、缺少
汇率/应回状态、同一资金记录重复归单、封存后被改写或工作簿不一致会阻止发布。

## 修改本技能后的验证

```powershell
python -X utf8 scripts/test_reconcile.py
python -X utf8 scripts/test_group_decisions.py
python -X utf8 scripts/test_simple_ledger.py
python -X utf8 -m compileall -q scripts
python -X utf8 ../.system/skill-creator/scripts/quick_validate.py .
```
