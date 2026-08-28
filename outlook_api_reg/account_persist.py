"""账号落盘字段归一化与多文件合并（注册 / 救援共用）。"""
from __future__ import annotations

from typing import Any

from .models import RegisterResult

_RECOVERY_KEYS = (
    "recovery_email",
    "recovery_password",
    "combo_recovery",
    "proofs_method",
    "proofs_satisfied",
)
_TOKEN_KEYS = (
    "refresh_token",
    "client_id",
    "combo",
    "login_refresh_token",
    "login_client_id",
    "combo_dual",
)
_RESCUE_KEYS = (
    "rescue_count",
    "last_rescue_at",
    "last_rescue_ok",
    "last_rescue_reason",
    "rescued_at",
    "rescued_scope",
)


def recovery_from_mapping(data: dict[str, Any]) -> tuple[str, str]:
    """从 post_login / proof_meta 等 dict 提取恢复邮箱字段。"""
    rec_e = str(data.get("recovery_email") or "").strip()
    rec_p = str(data.get("recovery_password") or "").strip()
    if rec_e and not rec_p:
        method = str(data.get("proofs_method") or "")
        if method == "cf_domain_recovery":
            rec_p = "cf_domain"
        elif method == "coolhs_mail_recovery":
            rec_p = "coolhs_mail"
    return rec_e, rec_p


def enrich_register_result(result: RegisterResult) -> RegisterResult:
    """把 post_login / extra 里的 proofs 信息补到 RegisterResult 顶层，避免落盘遗漏。"""
    post = (result.extra or {}).get("post_login") or {}
    if isinstance(post, dict):
        rec_e, rec_p = recovery_from_mapping(post)
        if rec_e and not result.recovery_email:
            result.recovery_email = rec_e
        if rec_p and not result.recovery_password:
            result.recovery_password = rec_p
        for key in ("proofs_method", "proofs_satisfied"):
            if post.get(key) and not result.extra.get(key):
                result.extra[key] = post[key]

    if not result.recovery_email:
        rec_e, rec_p = recovery_from_mapping(result.extra or {})
        if rec_e:
            result.recovery_email = rec_e
            if rec_p:
                result.recovery_password = rec_p

    if result.recovery_email and not result.recovery_password:
        method = str((result.extra or {}).get("proofs_method") or "")
        if method == "cf_domain_recovery":
            result.recovery_password = "cf_domain"
        elif method == "coolhs_mail_recovery":
            result.recovery_password = "coolhs_mail"
    return result


def merge_account_row(row: dict[str, Any], data: dict[str, Any], *, source: str = "") -> None:
    """同一邮箱多份 JSON 时合并，优先保留非空 recovery / token 字段。"""
    for key in _RECOVERY_KEYS + _TOKEN_KEYS + _RESCUE_KEYS:
        if data.get(key) and not row.get(key):
            row[key] = data[key]
    if data.get("rescue_count") is not None:
        row["rescue_count"] = max(int(row.get("rescue_count") or 0), int(data.get("rescue_count") or 0))
    if data.get("refresh_token") and not row.get("has_token"):
        row["has_token"] = True
    if data.get("combo_dual") or data.get("login_refresh_token"):
        row["login_token"] = True
    if row.get("recovery_email") and row.get("recovery_password"):
        row["has_recovery"] = True
    elif data.get("combo_recovery"):
        row["has_recovery"] = True
    for key in ("updated_at", "last_alive_at", "rescued_at"):
        if data.get(key) and (not row.get(key) or str(data[key]) > str(row.get(key) or "")):
            row[key] = data[key]
    if source:
        row["source"] = source
