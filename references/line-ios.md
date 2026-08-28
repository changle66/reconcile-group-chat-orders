# LINE iOS 备份读取

`extract_line_ios.py` 只读未加密的 Finder/iTunes iOS 备份，通过 `Manifest.db` 定位 LINE 数据库与媒体。
不要凭备份中的哈希文件名猜测媒体属于哪条消息。备份加密或核心数据库缺失时停止，并向用户说明缺少什么。

## 选择群聊

先列出备份中的群：

```powershell
python scripts/extract_line_ios.py <backup> --list-groups
```

再用明确的群 ID 或群名正则提取：

```powershell
python scripts/extract_line_ios.py <backup> --group-id <LINE群MID> -o <work>/line-normalized.json --timezone Asia/Bangkok --roster config/roster.yaml --force
```

也可重复使用 `--group-pattern` 或 `--group-id`。除非用户明确要求，不要使用 `--all-group-chats`，也不要仅凭
群名含“财务”等宽泛词语就纳入。

## 时间与身份

- LINE 消息时间是绝对时间，输出统一转换为 `Asia/Bangkok`；午夜附近的日期可能随时区转换变化。
- Excel 使用聊天消息时间。截图状态栏时间和图内交易时间不能替代消息时间。
- 文本、图片、相册图片和系统消息按其 LINE 内容类型读取；相册容器本身不生成媒体事件，每张相册图片单独处理。
- 加群、退群、移除、撤回等系统消息保留为上下文，但不记作资金流水。
- 有发送者 MID 时按 MID/显示名与 roster 判断角色。只有明确属于本人发出的普通消息，才可使用 `--self-name`
  补充显示名；不能把 sender 为空一概当作内部人员。

## 媒体

- 优先使用完整 attachment；只有缩略图时保留缩略图；两者都没有则标记 `availability=missing`。
- 可用媒体记录普通路径、SHA-256 和字节数。默认复制到输出旁的媒体目录；`--no-copy-media` 只在明确需要
  直接引用备份文件时使用。
- 语音、视频、贴纸、表情包和动图不作为资金证据，不需要 OCR 或转录。
- 缺失图片不能用相邻聊天金额、公式或旧表格反推。创建事件时将其写为 `missing_evidence_media`，使缺图在
  当前录单任务中保持可见。
