# LINE iOS 备份读取

统一流程使用 `reconcile.py start` 只读未加密的 Finder/iTunes iOS 备份，通过 `Manifest.db` 定位 LINE
数据库和媒体。不要按备份中的哈希文件名猜测消息归属。备份加密或核心数据库缺失时停止并明确报错。

## 在统一流程中使用

把包含 `Manifest.db` 的备份目录或其上级目录作为输入，`start` 会递归发现；也可显式重复传入
`--line-backup`：

```powershell
python scripts/reconcile.py start <原始导出根目录> --line-backup <LINE备份目录> --work <全新工作目录> --contains 小额
```

大额群单日任务改用：

```powershell
python scripts/reconcile.py start <原始导出根目录> --line-backup <LINE备份目录> --work <全新工作目录> --mode large --date YYYY-MM-DD
```

财务资料模式继续使用同一个 `Manifest.db` 读取器，只筛选实际群名包含 `财务资料群` 的群：

```powershell
python scripts/reconcile.py start <原始导出根目录> --line-backup <LINE备份目录> --work <全新工作目录> --mode finance
```

大额模式从备份读取全部群，再统一排除群名包含 `小额出` 或 `财务资料群` 的群。

统一流程按 LINE 数据库中解析出的实际群名筛选，不把显示名相同的 Telegram、WhatsApp 或 LINE 群混在
一起。新任务不要先运行旧提取脚本再拼接 normalized 文件，也不要使用 `--force` 覆盖历史输出。

需要诊断备份内容时，可以只读列群：

```powershell
python scripts/extract_line_ios.py <backup> --list-groups
```

`extract_line_ios.py` 的其他参数保留为内部兼容和故障诊断接口，不是新录单任务的公开工作流。

## 时间与身份

- LINE 绝对时间统一转换为 `Asia/Bangkok`。Excel 使用聊天消息时间，截图状态栏或图内交易时间不能替代。
- 文本、图片、相册图片和系统消息按 LINE 内容类型读取；相册中的每张图片单独进入媒体队列。
- 加群、退群、移除和撤回等系统消息只保留上下文，不记作资金流水。
- 有发送者 MID 时，按 MID、显示名和 roster 判断角色。
- sender 为空不能一概当作内部人员。只有 `ZSENDSTATUS` 明确证明是本机发出的普通消息时，才可按 self
  处理；系统状态或无法证明方向的记录保持 unknown。
- `--line-self-name` 只补充已经由发送状态证明为本机发出的显示名，不负责猜测身份。

## 媒体与哈希

- 优先使用完整 attachment；只有缩略图时保留缩略图；两者都没有则标记 `availability=missing`。
- `start` 可以把 LINE 媒体复制到本次工作目录，以保证路径稳定，但提取时不逐张计算媒体 SHA-256。
- LINE 数据库、`Manifest.db` 等来源文件在 `start` 时计入本次来源指纹。
- 订单模式只有审阅后判为 `fund` 的原图在 `review check/seal` 时计算证据哈希；财务资料模式对判为
  `document` 或 `chat_profile` 的原图采用相同保护。`finish` 再核对；参考图片不哈希。
- 语音、视频、贴纸、表情包和动图不作为资金证据，不需要 OCR 或转录。
- 缺失图片不能用相邻聊天金额、公式或旧表格反推。脚本会把它保留为可见的待确认记录。
