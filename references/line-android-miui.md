# LINE Android 小米/MIUI 备份读取

统一流程可直接读取小米或 MIUI 生成的 LINE Android 应用备份目录。典型目录同时包含
`descript.xml`、`LINE(jp.naver.line.android).bak` 和 `LINE(jp.naver.line.android).split`。聊天数据库与媒体在
`.bak` 内；`.split` 是拆分 APK 数据，不要拼接到 `.bak`，也不要把它当聊天附件。

## 在统一流程中使用

直接把备份目录或 `.bak` 文件作为 `start` 输入：

```powershell
python scripts/reconcile.py start <MIUI备份目录> --work <全新工作目录> --contains 小额
```

`start` 会递归发现受支持的 LINE Android `.bak`。也可显式重复传入：

```powershell
python scripts/reconcile.py start <原始导出根目录> --line-android-backup <目录或.bak> --work <全新工作目录> --contains 小额
```

大额群单日任务改用：

```powershell
python scripts/reconcile.py start <原始导出根目录> --line-android-backup <目录或.bak> --work <全新工作目录> --mode large --date YYYY-MM-DD
```

财务资料模式直接使用同一 MIUI 备份，只筛选实际群名包含 `财务资料群` 的群：

```powershell
python scripts/reconcile.py start <原始导出根目录> --line-android-backup <目录或.bak> --work <全新工作目录> --mode finance
```

大额模式从备份读取全部群，再统一排除群名包含 `小额出` 或 `财务资料群` 的群。

需要诊断备份内容时，只读列群：

```powershell
python scripts/extract_line_android_miui.py <目录或.bak> --list-groups
```

## 支持边界

- 读取 `ANDROID BACKUP` 头中 `compression=0`、`encryption=none` 的备份；压缩或加密备份会明确停止，不猜密钥、
  不尝试破密。
- 读取 `naver_line`、`contact` 及其 WAL/SHM 边车文件。数据库缺表、缺关键列或 TAR 结构损坏时停止并报告。
- 群名来自数据库，不从目录名或媒体路径猜测；相同 MID 的重复快照仍按统一流程拒绝合并。
- `.bak` 和同目录 `descript.xml` 进入来源指纹；`.split` 只含应用安装包，不参与聊天快照。

## 时间、身份与消息

- Android `created_time` 作为 Unix 绝对时间转换为 `Asia/Bangkok`。
- 发送者 MID 通过 `contact` 数据库映射本地覆盖名或资料名，再由 `config/roster.yaml` 判定内部人员；无法映射时
  保留 MID 并把角色交给正常未知流程，不把空发送者默认当内部人员。
- 回复关系使用消息参数中的服务端消息 ID；无法在本次备份找到原消息时不虚构引用。
- LINE 加群、退群、撤回、贴纸、语音、视频及其他上下文消息保留在时间线但排除资金核算。

## 媒体

- 图片按数据库本地消息 ID 对应到 `ef/chats/<群MID>/messages/<本地消息ID>`，依次优先使用 `.original`、无后缀
  主文件、`.thumb`。
- `start` 只把选中群的图片解出到本次工作目录；不使用 TAR 的原始路径作为可变外部引用，也不在提取阶段逐图
  计算 SHA-256。
- 只有缩略图时保留并警告；所有候选都缺失时写 `availability=missing/missing_kind=missing_from_backup`，不能按
  相邻聊天金额或旧台账补造凭证。
- 图片签名用于生成可查看的扩展名；未知图片签名仍以原始字节保留，不执行 OCR 或格式转换。
