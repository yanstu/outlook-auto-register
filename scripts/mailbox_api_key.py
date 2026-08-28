#!/usr/bin/env python3
"""Mailbox API 的 key 管理：建 key、看 key、停用 key。

用法：
  python3 scripts/mailbox_api_key.py create --name kimi-register \
      --scopes mailboxes:read,fields:basic,messages:read,otp:read
  python3 scripts/mailbox_api_key.py create --name single-box \
      --scopes fields:basic,otp:read --grant a@outlook.com --grant b@outlook.com
  python3 scripts/mailbox_api_key.py list
  python3 scripts/mailbox_api_key.py revoke <key_id>
  python3 scripts/mailbox_api_key.py audit --limit 20

新建的 key 只在创建那一刻打印一次，库里只留 pbkdf2 摘要，丢了只能重建。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from outlook_api_reg.mailbox_gateway import store  # noqa: E402

DEFAULT_SCOPES = "mailboxes:read,fields:basic,messages:read,otp:read"


def cmd_create(args: argparse.Namespace) -> int:
    scopes = store.normalize_scopes(args.scopes)
    if not scopes:
        print("至少要给一个 scope，例如 --scopes " + DEFAULT_SCOPES)
        return 2
    created = store.create_service_key(
        args.name, scopes, grants=args.grant, expires_at=args.expires_at
    )
    print("id     :", created["id"])
    print("name   :", created["name"])
    print("scopes :", ",".join(created["scopes"]))
    print("邮箱范围:", ", ".join(created["grants"]) if created["grants"] else "全部邮箱")
    print("key    :", created["key"])
    print("\n这把 key 只显示这一次，请立刻存好。")
    return 0


def cmd_list(_args: argparse.Namespace) -> int:
    items = store.list_principals()
    if not items:
        print("还没有任何 key。")
        return 0
    for item in items:
        scope_text = ",".join(item["scopes"]) or "-"
        grants = [g for g in item.get("grants") or [] if g]
        scope_of = ", ".join(grants) if grants else "全部邮箱"
        state = "启用" if item["enabled"] else "已停用"
        print(f"{item['id']}  {state}  {item['kind']:<7} {item['name']}")
        print(f"    scopes={scope_text}  邮箱范围={scope_of}  建于={item['created_at']}")
    return 0


def cmd_revoke(args: argparse.Namespace) -> int:
    if store.revoke_principal(args.key_id):
        print(f"{args.key_id} 已停用，名下会话一并清掉。")
        return 0
    print(f"没有这把 key: {args.key_id}")
    return 1


def cmd_audit(args: argparse.Namespace) -> int:
    rows = store.recent_audit(args.limit)
    if not rows:
        print("还没有访问记录。")
        return 0
    for r in rows:
        print(f"{r['ts']}  {r['status']:>3}  {r['method']:<6} {r['path']}"
              f"  {r['email'] or '-'}  {r['principal_id'] or '-'}  {r['detail']}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Mailbox API key 管理")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_create = sub.add_parser("create", help="新建一把 service key")
    p_create.add_argument("--name", default="", help="给这把 key 起个名字（调用方）")
    p_create.add_argument("--scopes", default=DEFAULT_SCOPES, help="逗号分隔")
    p_create.add_argument("--grant", action="append", default=None,
                          help="限定可访问的邮箱，可重复；不给 = 全部邮箱")
    p_create.add_argument("--expires-at", dest="expires_at", default="",
                          help="到期时间（ISO，如 2027-01-01T00:00:00）")
    p_create.set_defaults(func=cmd_create)

    p_list = sub.add_parser("list", help="列出所有 key")
    p_list.set_defaults(func=cmd_list)

    p_revoke = sub.add_parser("revoke", help="停用一把 key")
    p_revoke.add_argument("key_id")
    p_revoke.set_defaults(func=cmd_revoke)

    p_audit = sub.add_parser("audit", help="看最近的访问记录")
    p_audit.add_argument("--limit", type=int, default=50)
    p_audit.set_defaults(func=cmd_audit)

    args = ap.parse_args()
    return int(args.func(args))


if __name__ == "__main__":
    sys.exit(main())
