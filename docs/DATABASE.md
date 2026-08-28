# SQLite 统一存储（部署用）

所有业务数据写入 **`accounts/outlook.db`**（可用 `OUTLOOK_DB_PATH` 覆盖）。

## 表结构

| 表 | 用途 |
|---|---|
| `accounts` | 账号主表（邮箱 PK、密码、token、combo、批次、救援统计） |
| `account_meta` | 备注、标签、测活缓存 |
| `register_jobs` | 注册批次历史 |
| `proxies` / `proxy_bindings` / `proxy_settings` | 代理池（模板、sticky 绑定、分配策略） |
| `proxy_events` | 代理使用事件（代理商×代理国×注册国成功率统计） |
| `rescue_events` | 重登事件日志 |
| `api_principals` | Mailbox API 的调用方（service key / 用户端），只存 pbkdf2 摘要 |
| `api_grants` | 某个调用方能看哪些邮箱（email 为空 = 全部）与 scope 覆盖 |
| `api_sessions` | 用户端登录会话（绑定单个邮箱、带有效期） |
| `api_audit` | Mailbox API 访问记录（时间、路径、邮箱、状态），不含令牌 |
| `app_meta` | 迁移标记、最近备份时间 |

`api_*` 四张表随 schema v4 引入，老库启动时自动补齐，见 [MAILBOX_API.md](MAILBOX_API.md)。

## 首次启动

1. 自动建表（WAL 模式）
2. 若存在遗留 `accounts/*.json`、`webapp_meta.json`、`webapp_jobs.json`、combo txt → **迁入 SQLite**
3. 迁移前自动热备份到 `accounts/backups/outlook_*_pre-migrate.db`
4. 源 JSON 移至 `accounts/.archive/<时间戳>/`

## 备份

```bash
# CLI
python scripts/backup_db.py --tag nightly

# API
POST /api/database/backup
GET  /api/database
POST /api/database/migrate   # 手动重跑遗留迁移
```

默认保留最近 **30** 份备份。

## 导出

账号 combo 仍通过 Web「导出」或 `GET /api/accounts/export` 从数据库生成，**不再**依赖 `accounts.txt` 作为主存储。
