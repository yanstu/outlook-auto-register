# 长存账号流水线（孵化 → 保活 → 分层导出）

面向自有 Outlook 账号的生命周期管理：新号先静默孵化，定期轻量保活轮换
`refresh_token`，达标后再导出六段 `combo_recovery`。

## 总览

```
注册(graph_recovery 默认)
    ↓
写入 SQLite + tags=incubating
    ↓  OUTLOOK_INCUBATION_HOURS（默认 48h）
孵化期：批量测活 / 定时保活跳过（不打微软）
    ↓
出孵化 → cron 每 2 天 keepalive_from_db
    ↓  存活 ≥ min_days（默认 7）
export_long_lived / GET /api/export/long-lived
    → accounts/long_lived.txt
```

## 1. 默认产出：`graph_recovery`

- Web「产出格式」与 `POST /api/register` 在省略 `token_mode` 时默认
  **`graph_recovery`**（Graph 四段 + 恢复邮箱 = 六段）。
- 导出下拉默认 **`recovery`**（六段）。
- `/api/config` 字段：`default_token_mode=graph_recovery`、
  `default_export_format=recovery`、`incubation_hours`。

## 2. 孵化期

| 项 | 说明 |
| --- | --- |
| 环境变量 | `OUTLOOK_INCUBATION_HOURS`（默认 `48`） |
| 新号 | `save_register_result` 写入 meta 标签 `incubating` |
| API | `/api/accounts` 每条带 `incubating: bool`、`incubation_until: ISO` |
| 批量测活 | `POST /api/verify-batch` 对孵化号返回 `skipped: true`，**不调用**微软 |
| 保活 | Web `/api/keepalive` 与 `scripts/keepalive_from_db.py` 同样跳过 |

实现：`outlook_api_reg/lifecycle.py`（按 `created_at + hours` 动态判断，过期自动清标签）。

## 3. 定时保活

```bash
# 手动
python3 scripts/keepalive_from_db.py [--proxy host:port:user:pass] [--concurrency 5]

# 或
./scripts/cron_keepalive.sh

# 安装 crontab（每 2 天 03:17）
./scripts/install_keepalive_cron.sh
```

- 读 SQLite 账号，跳过孵化期，调用 `keepalive_one` 轮换 token 并写回库。
- 日志：`accounts/keepalive.log`。

## 4. 分层导出

```bash
python3 scripts/export_long_lived.py --min-days 7
# → accounts/long_lived.txt（combo_recovery 行）
```

HTTP：

```http
GET /api/export/long-lived?min_days=7
```

条件：非孵化、年龄 ≥ `min_days`、有可用 `combo_recovery`（或可拼装的 recovery 字段）。

## 5. 测试

```bash
python3 -m unittest tests.test_lifecycle -v
```

## 相关文件

- `outlook_api_reg/lifecycle.py`
- `outlook_api_reg/account_store.py`
- `webapp/server.py` / `webapp/static/index.html`
- `scripts/keepalive_from_db.py` / `cron_keepalive.sh` / `install_keepalive_cron.sh`
- `scripts/export_long_lived.py`
- `tests/test_lifecycle.py`
