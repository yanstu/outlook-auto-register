# 外部账号导入（自有资产合并）

同一台 38 服务器上还跑着 qoderji（另一套系统，同属 Hugo）。qoderji 的
`email_inventory` 表里攒了几千个早先已经注册好的 Outlook 邮箱（`email----password
----client_id----refresh_token` 四段 combo；`password` 段常是占位符 `x`，真正能
用的是 OAuth `refresh_token`）。这些号本来就是自己的资产，只是散在另一套库里；
本页说明怎么把它们并进 outlook-auto-register 的账号库，跟本地注册的号一起走同一套
运维台 / Mailbox API 管理、收信。

只读合并：任何导入路径都不会写回 qoderji 的数据库，也不会影响它自己的租约 /
库存统计。

## 核心概念

### combo 解析（`outlook_api_reg/external_import.py`）

- **4 段**：`email----password----client_id----refresh_token`（Graph 四段，qoderji
  的标准形态）。
- **6 段**：末两段固定视为登录令牌 `login_client_id----login_refresh_token`，语义
  与运维台既有的粘贴导入完全一致。
- 缺 `refresh_token` 的行会被判为无效（`invalid`）——没有令牌就没法收信，导进来
  也用不了。
- 按邮箱去重（不分大小写）：本地账号库里已存在的邮箱直接跳过，**不覆盖任何已有
  字段**；同一批次内重复邮箱只取第一条。

### 孵化期（`skip_incubation`）

本地新注册的号默认进入 `OUTLOOK_INCUBATION_HOURS`（默认 48h）孵化期，期间批量测活
/ 保活跳过，避免对新号高频打微软接口触发风控（见 [`docs/LONG_LIVED_PIPELINE.md`](LONG_LIVED_PIPELINE.md)）。

外部导入的号早就存活过一段时间，不需要（也不应该）再走这段冷启动：

- `skip_incubation=true`（qoderji 导入默认如此）：`created_at` 回填到
  `OUTLOOK_EXTERNAL_IMPORT_BACKDATE_DAYS`（默认 **30** 天）前，账号立即可用，运维台
  列表里不带 `incubating` 标签。
- `skip_incubation=false`（粘贴导入的默认值，兼容老行为）：`created_at` 记为当前
  时间，仍受 48h 孵化期限制。

### 为什么默认排除 `status=dead`

qoderji 给邮箱标记的状态含义：

| qoderji `status` | 含义 | 默认是否导入 |
|---|---|---|
| `untried` | 还没被 qoderji 用掉的原料 | 导入 |
| `leased` | 正被某台注册机租着（只读不受影响） | 导入 |
| `consumed` | 已经被 qoderji 用去注册过一个产品账号，但邮箱本身仍能收信 | 导入 |
| `dead` | OAuth 永久失效（如 `invalid_grant` / 账号被判定异常） | **默认排除** |

`dead` 的号连微软那边的令牌都已经作废，导进来也收不了信，所以默认不导入。可以用
`--status dead` 明确要它（不建议）。

## Web API

### 增强：`POST /api/accounts/import`

粘贴 combo 文本导入，行为对未传新参数的调用完全兼容。

```json
{
  "text": "a@outlook.com----pwd----cid----rt\n...",
  "source": "manual-merge",
  "batch_label": "手工合并",
  "skip_incubation": false
}
```

| 字段 | 默认 | 说明 |
|---|---|---|
| `text` | `""` | 必填，每行一条 combo |
| `source` | `""` | 来源标记，写入 `legacy_source` 与 `account_meta.tags`（`src:<source>`） |
| `batch_label` | `""` | 运维台批次筛选用的标签；留空时用 `外源导入:<source>` |
| `skip_incubation` | `false` | 是否跳过孵化期（回填 `created_at`） |

返回：`{ok, imported, duplicate, invalid, six_seg, source, batch_label,
skip_incubation}`。

### 新增：`POST /api/accounts/import/qoderji`

从 qoderji 的 `email_inventory` 拉取并导入。只读连接（`sqlite3` URI `mode=ro`），
不会跟 qoderji 自己的写操作抢锁。

```json
{
  "db_path": null,
  "statuses": null,
  "batch_id": null,
  "limit": null,
  "batch_label": "",
  "skip_incubation": true,
  "dry_run": false
}
```

| 字段 | 默认 | 说明 |
|---|---|---|
| `db_path` | `null` | 显式指定 sqlite 路径（可逗号分隔多个/带 glob）；不传则走 `QODERJI_EMAIL_DB` 环境变量，再不然按 `/opt/qoderji/data/*.db`、`/opt/qoderji/*.db` 探测 |
| `statuses` | `null` | 覆盖状态过滤，不传则用 `untried,leased,consumed`（排除 `dead`） |
| `batch_id` | `null` | 只拉 qoderji 某一个 `batch_id` |
| `limit` | `null` | 最多拉取 / 导入多少条 |
| `batch_label` | `""` | 留空时用 `qoderji导入` |
| `skip_incubation` | `true` | 外源号默认跳过孵化期 |
| `dry_run` | `false` | 只统计不写库，用来先看看有多少条能导 |

返回除 `imported/duplicate/invalid/six_seg` 外，还带 `db_files`（实际扫到的库文件）、
`fetch_stats`（每个库的扫描/解析/按状态计数）、`fetched`（总共取到多少条待导入）。
找不到任何库文件时返回 `404`。

示例：

```bash
BASE=https://outlook.coolhs.com
# 先 dry-run 看看有多少能导
curl -s -X POST "$BASE/api/accounts/import/qoderji" \
  -H "X-Console-Password: $OUTLOOK_CONSOLE_PASSWORD" -H "Content-Type: application/json" \
  -d '{"dry_run": true}'
# 确认没问题再真导
curl -s -X POST "$BASE/api/accounts/import/qoderji" \
  -H "X-Console-Password: $OUTLOOK_CONSOLE_PASSWORD" -H "Content-Type: application/json" \
  -d '{}'
```

## 运维台 UI

「导入账号」弹窗底部多了一个「从 qoderji 拉取」入口：填一个可选的数量上限，点一下
就直接拉取并导入（默认状态过滤、默认跳过孵化期），不用先导出成文本再粘贴。

## CLI：`scripts/import_external_outlook.py`

不想通过 HTTP、或要在没起 web 服务的情况下批量跑，可以直接用 CLI：

```bash
# 从一份 4/6 段 combo 文本文件导入
.venv/bin/python scripts/import_external_outlook.py --file accounts.txt --source manual

# 先 dry-run 看看 qoderji 那边能拉多少
.venv/bin/python scripts/import_external_outlook.py --qoderji --dry-run

# 真导入，限定状态、限定数量
.venv/bin/python scripts/import_external_outlook.py --qoderji \
  --status consumed --status untried --limit 500

# 显式指定库路径（默认路径探测失败、或要指定某个历史拷贝时用）
.venv/bin/python scripts/import_external_outlook.py --qoderji \
  --qoderji-db /opt/qoderji/data/cards.db
```

参数与 Web API 字段一一对应；`--file` 和 `--qoderji` 二选一，`--no-skip-incubation`
关掉默认的孵化期跳过。全部输出 JSON，方便接脚本或人工核对。

## qoderji `email_inventory` 表结构（供对照）

来自 `qoderji_server/email_pool.py::init_db()`（38 服务器实测，实际数据库是
`/opt/qoderji/data/cards.db`，`email_inventory` 只是其中一张表）：

```sql
CREATE TABLE IF NOT EXISTS email_inventory (
    email        TEXT PRIMARY KEY,          -- 小写归一，做去重主键
    raw          TEXT NOT NULL DEFAULT '',  -- 原始行(email----x----guid----M.C...)
    batch_id     TEXT NOT NULL DEFAULT '',
    source       TEXT NOT NULL DEFAULT '',  -- 订单文件名/来源
    status       TEXT NOT NULL DEFAULT 'untried',  -- untried|leased|consumed|dead
    leased_by    TEXT NOT NULL DEFAULT '',
    leased_name  TEXT NOT NULL DEFAULT '',
    lease_job    TEXT NOT NULL DEFAULT '',
    lease_expires_at REAL,
    dead_reason  TEXT,
    added_at     REAL NOT NULL,
    leased_at    REAL,
    consumed_at  REAL
);
CREATE INDEX IF NOT EXISTS idx_ei_status ON email_inventory(status);
CREATE INDEX IF NOT EXISTS idx_ei_batch  ON email_inventory(batch_id);
CREATE INDEX IF NOT EXISTS idx_ei_lease  ON email_inventory(lease_expires_at);
```

实测规模（2026-08-28）：`consumed` 9916、`dead` 1058、`untried` 2，`raw` 清一色 4 段
`----` combo，域名清一色 `@outlook.com`。
