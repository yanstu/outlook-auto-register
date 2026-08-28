# Mailbox API v1

给其他注册项目用的收信接口：拿一枚 key，就能通过 HTTP 列邮箱、读信、取验证码、导出账号字段，
不用碰这边的数据库，也不用自己维护 Outlook 的 refresh token。

- 根路径：`/api/v1`（和运维台同一个服务，默认 `http://127.0.0.1:8890`）
- 所有响应都是 JSON；出错时是 `{"code": "...", "message": "..."}`，可能带额外字段
- 时间一律 ISO 字符串（微软返回的收信时间带 `Z`）

## 快速上手

```bash
BASE=http://127.0.0.1:8890
KEY=mbx_sk_xxxxxxxxxxxx_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

# 1. 有哪些邮箱
curl -s -H "Authorization: Bearer $KEY" "$BASE/api/v1/mailboxes?limit=5"

# 2. 拿其中一个的 mailbox_id 去取验证码（最多等 60 秒）
curl -s -H "Authorization: Bearer $KEY" \
  "$BASE/api/v1/mailboxes/$MAILBOX_ID/otp?wait_seconds=60"
```

## 鉴权

统一用 `Authorization: Bearer <token>`。有两种 token：

| 类型 | 长什么样 | 给谁用 | 能看到什么 |
| --- | --- | --- | --- |
| service key | `mbx_sk_<id>_<secret>` | 其他项目、脚本、定时任务 | 按 scopes 和授权邮箱集合决定 |
| 用户会话 | `mbx_sess_<id>_<secret>` | 未来的用户端页面 | **只有登录的那一个邮箱** |

### 建一把 service key

服务端 CLI（推荐，key 不经过网络）：

```bash
python3 scripts/mailbox_api_key.py create \
  --name kimi-register \
  --scopes mailboxes:read,fields:basic,messages:read,otp:read
```

只想放开固定几个邮箱就加 `--grant`（可重复），不加 `--grant` 等于放开全部邮箱：

```bash
python3 scripts/mailbox_api_key.py create --name one-box \
  --scopes fields:basic,otp:read --grant abc@outlook.com
```

也可以用接口建（调用方自己得先有一枚带 `admin` 的 key）：

```bash
curl -s -X POST "$BASE/api/v1/auth/keys" \
  -H "Authorization: Bearer $ADMIN_KEY" -H 'Content-Type: application/json' \
  -d '{"name":"kimi-register","scopes":"mailboxes:read,otp:read","grants":["abc@outlook.com"]}'
```

**key 只在创建那一刻返回一次**，库里只留 pbkdf2-sha256 摘要，丢了只能重建。
不用了就停：`python3 scripts/mailbox_api_key.py revoke <id>`，或
`DELETE /api/v1/auth/keys/{id}`。停用会同时清掉这把 key 名下的所有会话。

第一枚 admin key 可以直接写在环境变量里，服务启动时自动写库：

```
OUTLOOK_MAILBOX_API_ADMIN_KEY=随便一串足够长的随机字符串
```

## scopes

| scope | 放开什么 |
| --- | --- |
| `mailboxes:read` | 列邮箱、看邮箱详情 |
| `fields:basic` | 邮箱地址、批次、孵化状态、能否读信、创建时间 |
| `fields:sensitive` | 密码、恢复邮箱、client_id、combo 文本（只发给 service key） |
| `messages:read` | 列信、读单封信 |
| `otp:read` | 长轮询取验证码 |
| `admin` | 建 key、看 key、停 key |
| `incubation:bypass` | 无视孵化期直接读信（一般只给 admin） |

`*` 表示全部 scope，环境变量写入的 admin key 拿到的是 `admin,*`。
授权到具体邮箱时还可以再收窄（`api_grants.scopes_override`），
最终生效的是「key 的 scopes ∩ 这个邮箱上的收窄结果」。

## 端点

### 鉴权类

#### `POST /api/v1/auth/login`

用户端登录，邮箱 + 邮箱密码换一枚只能看自己这一个邮箱的会话令牌。

```bash
curl -s -X POST "$BASE/api/v1/auth/login" -H 'Content-Type: application/json' \
  -d '{"email":"abc@outlook.com","password":"邮箱密码"}'
```

```json
{
  "token": "mbx_sess_1a2b3c4d5e6f_...",
  "expires_in": 86400,
  "expires_at": "2026-08-29T09:00:00",
  "mailbox_id": "mbx_YWJjQG91dGxvb2suY29t",
  "email": "abc@outlook.com",
  "scopes": ["fields:basic", "messages:read", "otp:read"]
}
```

会话有效期由 `OUTLOOK_MAILBOX_API_SESSION_HOURS` 决定（默认 24 小时）。
会话拿不到 `fields:sensitive`，所以 `profile=full` / `profile=combo` 对它是 403。

#### `POST /api/v1/auth/keys` · `GET /api/v1/auth/keys` · `DELETE /api/v1/auth/keys/{id}`

建 / 看 / 停 service key，都要 `admin`。

#### `GET /api/v1/auth/me`

当前 token 是谁、有哪些 scope、能看哪些邮箱。接调用方时先打这个自检最省事。

### 邮箱类

#### `GET /api/v1/mailboxes`

分页列表。参数：`limit`（1–500，默认 50）、`offset`、`q`（邮箱模糊匹配）、
`batch`（批次标签）、`readable_only`。

```json
{
  "total": 128,
  "limit": 50,
  "offset": 0,
  "items": [
    {
      "mailbox_id": "mbx_YWJjQG91dGxvb2suY29t",
      "email": "abc@outlook.com",
      "batch": "B7",
      "created_at": "2026-08-01T10:00:00",
      "incubating": false,
      "readable": true,
      "has_token": true,
      "tags": []
    }
  ]
}
```

字段随 scope 增减：没有 `fields:sensitive` 就看不到 `password` / `recovery_email` /
`client_id`。**refresh token 在任何 scope 下都不会出现在响应里**，
要账号文本请走 `fields?profile=combo`。

#### `GET /api/v1/mailboxes/{mailbox_id}` · `GET /api/v1/mailboxes/by-email/{email}`

单个邮箱详情，两个入口内容一样。`mailbox_id` 是邮箱地址的 urlsafe base64，
前缀 `mbx_`，稳定可缓存；不想自己算就直接用 `by-email`。

#### `GET /api/v1/mailboxes/{mailbox_id}/fields?profile=basic|full|combo`

按用途取一组字段：

| profile | 内容 | 需要 |
| --- | --- | --- |
| `basic` | email、readable、incubating、batch、created_at | `fields:basic` |
| `full` | basic + password、recovery_email、recovery_password、client_id | `fields:sensitive` |
| `combo` | 四段 / 六段账号文本 | `fields:sensitive` |

```bash
curl -s -H "Authorization: Bearer $KEY" \
  "$BASE/api/v1/mailboxes/$MAILBOX_ID/fields?profile=combo"
```

```json
{
  "mailbox_id": "mbx_YWJjQG91dGxvb2suY29t",
  "profile": "combo",
  "combo": "abc@outlook.com----密码----client_id----refresh_token----恢复邮箱----恢复密码",
  "segments": 6
}
```

有恢复邮箱就给六段，没有就给四段，`segments` 直接告诉你是哪种。

### 读信类

#### `GET /api/v1/mailboxes/{mailbox_id}/messages`

参数：`folder`（`inbox` / `junkemail`，默认 `inbox`）、`limit`（1–50，默认 20）、
`after`（只要这个时间之后收到的）、`from`（发件人包含）、`subject_contains`、
`mode`（`auto` / `graph` / `outlook_rest`，默认 `auto`）。

```bash
curl -s -H "Authorization: Bearer $KEY" \
  "$BASE/api/v1/mailboxes/$MAILBOX_ID/messages?folder=junkemail&limit=10&subject_contains=code"
```

```json
{
  "email": "abc@outlook.com",
  "folder": "junkemail",
  "mode": "graph",
  "count": 1,
  "messages": [
    {
      "id": "AAMkAGI...",
      "folder": "junkemail",
      "subject": "Microsoft account security code",
      "from": "account-security-noreply@accountprotection.microsoft.com",
      "received": "2026-08-28T09:00:00Z",
      "preview": "Security code: 123456",
      "body": "<html>…</html>"
    }
  ]
}
```

#### `GET /api/v1/mailboxes/{mailbox_id}/messages/{message_id}`

单封信。`message_id` 用上一步返回的 `id`。只在最近的收件箱和垃圾邮件里找，
太老的信找不到会返回 `message_not_found`。

#### `GET /api/v1/mailboxes/{mailbox_id}/otp`

长轮询取验证码：命中就立刻返回，没命中就在超时前反复看收件箱和垃圾邮件。

参数：`wait_seconds`（0–60，默认 60）、`after`、`sender`、`subject_contains`、`mode`。

```bash
# 触发发码之前先记下时间，用 after 挡掉旧信
NOW=$(date -u +%Y-%m-%dT%H:%M:%SZ)
curl -s -H "Authorization: Bearer $KEY" \
  "$BASE/api/v1/mailboxes/$MAILBOX_ID/otp?wait_seconds=60&after=$NOW"
```

```json
{
  "email": "abc@outlook.com",
  "found": true,
  "code": "123456",
  "mode": "graph",
  "message": { "id": "AAMkAGI...", "subject": "Microsoft account security code", "…": "…" }
}
```

不传 `sender` / `subject_contains` 时只认微软安全码那类邮件；
想收别家的码（比如你自己站点发的），传上 `sender` 或 `subject_contains`
就按你给的条件筛，不再套微软的判断。

没等到是正常返回，不是错误：

```json
{ "email": "abc@outlook.com", "found": false, "code": "", "mode": "graph", "waited_seconds": 60 }
```

一次最多等 60 秒。要等更久就自己再打一次，`after` 保持不变。

### 运维类

#### `GET /api/v1/health`

要带有效 token，返回 `{"ok": true, "readable_count": 42, "schema_version": 4}`。
`readable_count` 只统计这枚 token 能看到的邮箱。

## 错误

| HTTP | `code` | 什么情况 |
| --- | --- | --- |
| 400 | `bad_mailbox_id` / `bad_profile` / `bad_folder` / `bad_mode` | 参数不对 |
| 401 | `unauthorized` | 没带 token、token 过期或已停用 |
| 401 | `invalid_credentials` | 登录的邮箱或密码不对 |
| 403 | `scope_required` | token 缺这个 scope，响应里带 `required_scope` |
| 403 | `mailbox_forbidden` | 这枚 token 没被授权看这个邮箱 |
| 404 | `mailbox_not_found` | 库里没有这个邮箱 |
| 404 | `no_token` | 这个邮箱没有读信令牌，收不了信 |
| 404 | `message_not_found` | 最近的信里没有这个 id |
| 422 | — | 参数越界（如 `wait_seconds=600`），FastAPI 的校验响应 |
| 423 | `incubating` | 账号还在孵化期，响应里带 `until` |
| 502 | `token_dead` | 读信令牌换不出 access token，需要重登或救援 |
| 503 | `upstream_unavailable` | 微软那边网络抖动，稍后重试 |

孵化期是新号的保护窗口（`OUTLOOK_INCUBATION_HOURS`，默认 48 小时），
这期间读信和取码一律 423，避免刚建好的号被高频访问打上风控。
确实要提前读，就给那枚 key 加 `incubation:bypass`。

## 令牌与账号数据怎么处理

- 每次读信前先用 refresh token 换一次 access token，顺便判活。
  微软轮换出新的 refresh token 时会**自动写回账号库**（`combo` / `combo_dual` /
  `combo_recovery` 的第四段一起更新），调用方不用管。
- refresh token 不会出现在任何响应里。要账号文本只有 `fields?profile=combo` 一条路，
  而且要 `fields:sensitive`。
- 每次调用都会在 `api_audit` 里留一行：时间、调用方 id、方法、路径、邮箱、状态码。
  **不记任何令牌、密码、验证码**。看最近记录：
  `python3 scripts/mailbox_api_key.py audit --limit 20`。

## 环境变量

```
OUTLOOK_MAILBOX_API_ENABLED=1        # 0 = 不挂载 /api/v1，只留运维台
OUTLOOK_MAILBOX_API_ADMIN_KEY=       # 非空则首次启动写入一枚 admin key
OUTLOOK_MAILBOX_API_SESSION_HOURS=24 # 用户端会话有效期
```

## 后面往用户端扩

现在这版是「service key 能用、user session 已经能跑通」的状态。用户端页面要接的话：

- 登录已经有了：`POST /api/v1/auth/login` 拿 `mbx_sess_...`，
  之后所有请求带这枚 token，服务端强制只放行绑定的那一个邮箱，
  前端不需要也拿不到别人的数据。
- 会话默认能用 `fields:basic` / `messages:read` / `otp:read`，
  刚好够「看自己的邮箱、读自己的信、取自己的码」，敏感字段天然被挡在外面。
- 要加新能力（比如让用户自己改备注）就新增一个 scope，
  在 `api_principals.scopes` 里放开，不用动鉴权主干。
- 要给某个 service key 单独收窄某个邮箱的权限，写 `api_grants.scopes_override` 即可，
  取交集的逻辑已经在。

## 代码在哪

| 文件 | 管什么 |
| --- | --- |
| `outlook_api_reg/mailbox_gateway/routes.py` | 路由与参数校验 |
| `outlook_api_reg/mailbox_gateway/auth.py` | Bearer 解析、principal、scope 判定、访问记录 |
| `outlook_api_reg/mailbox_gateway/service.py` | mailbox_id、字段裁剪、读信取码、令牌写回 |
| `outlook_api_reg/mailbox_gateway/store.py` | principal / grant / session / audit 读写 |
| `outlook_api_reg/mailbox_gateway/schemas.py` | 请求与响应模型 |
| `scripts/mailbox_api_key.py` | key 的建 / 看 / 停 / 审计 |
