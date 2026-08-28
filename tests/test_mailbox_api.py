"""Mailbox API v1：字段裁剪、鉴权、以及跑通 TestClient 的端到端行为。"""
from __future__ import annotations

import os
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta
from unittest import mock

_TMP_DIR = tempfile.mkdtemp(prefix="mailbox-api-test-")
_PREV_ENV = {
    k: os.environ.get(k)
    for k in (
        "OUTLOOK_DB_PATH",
        "OUTLOOK_MAILBOX_API_ENABLED",
        "OUTLOOK_MAILBOX_API_ADMIN_KEY",
        "OUTLOOK_MAILBOX_API_SESSION_HOURS",
        "OUTLOOK_INCUBATION_HOURS",
    )
}
os.environ["OUTLOOK_DB_PATH"] = os.path.join(_TMP_DIR, "outlook.db")
os.environ["OUTLOOK_MAILBOX_API_ENABLED"] = "1"
os.environ["OUTLOOK_MAILBOX_API_ADMIN_KEY"] = "test-admin-key"
os.environ["OUTLOOK_MAILBOX_API_SESSION_HOURS"] = "2"
os.environ["OUTLOOK_INCUBATION_HOURS"] = "48"

from fastapi.testclient import TestClient  # noqa: E402

from outlook_api_reg import account_store, database as db, graph_mail  # noqa: E402
from outlook_api_reg.mailbox_gateway import auth, service, store  # noqa: E402
from outlook_api_reg.mailbox_gateway.errors import MailboxApiError  # noqa: E402

MATURE = "mature@outlook.com"
FRESH = "fresh@outlook.com"
OTHER = "other@outlook.com"
MATURE_PASSWORD = "MaturePass!234"

ADMIN_HEADERS = {"Authorization": "Bearer test-admin-key"}


def _iso(dt: datetime) -> str:
    return dt.isoformat(timespec="seconds")


def _seed_accounts() -> None:
    old = _iso(datetime.now() - timedelta(days=30))
    recent = _iso(datetime.now() - timedelta(hours=1))
    rows = [
        {
            "email": MATURE,
            "password": MATURE_PASSWORD,
            "client_id": "9e5f94bc-e8a4-4e73-b8be-63364c29d753",
            "refresh_token": "rt-mature",
            "recovery_email": "rec@example.com",
            "recovery_password": "recpwd",
            "combo": f"{MATURE}----{MATURE_PASSWORD}----cid----rt-mature",
            "combo_recovery": f"{MATURE}----{MATURE_PASSWORD}----cid----rt-mature----rec@example.com----recpwd",
            "batch_label": "B1",
            "created_at": old,
            "updated_at": old,
        },
        {
            "email": FRESH,
            "password": "FreshPass!234",
            "refresh_token": "rt-fresh",
            "combo": f"{FRESH}----FreshPass!234----cid----rt-fresh",
            "batch_label": "B2",
            "created_at": recent,
            "updated_at": recent,
        },
        {
            "email": OTHER,
            "password": "OtherPass!234",
            "refresh_token": "",
            "batch_label": "B1",
            "created_at": old,
            "updated_at": old,
        },
    ]
    conn = db.connect()
    try:
        for row in rows:
            account_store.upsert_account_dict(conn, row)
        account_store.upsert_meta_dict(conn, MATURE, {"tags": [], "verify": {"ok": True}})
        conn.commit()
    finally:
        conn.close()


def setUpModule() -> None:
    db._initialized = False
    db.ensure_initialized()
    conn = db.connect()
    try:
        db._app_set(conn, "legacy_files_migrated", "1")
        conn.commit()
    finally:
        conn.close()
    _seed_accounts()


def tearDownModule() -> None:
    for key, value in _PREV_ENV.items():
        if value is None:
            os.environ.pop(key, None)
        else:
            os.environ[key] = value
    db._initialized = False
    shutil.rmtree(_TMP_DIR, ignore_errors=True)


# ── 纯函数：id、scope、摘要、字段裁剪 ────────────────────────────────────────


class TestIdentifiers(unittest.TestCase):
    def test_mailbox_id_round_trip(self):
        mid = service.mailbox_id_for("Alice@Outlook.com")
        self.assertTrue(mid.startswith("mbx_"))
        self.assertNotIn("=", mid)
        self.assertEqual(service.email_from_mailbox_id(mid), "alice@outlook.com")

    def test_parse_mailbox_ref_accepts_email(self):
        self.assertEqual(service.parse_mailbox_ref("Bob@Outlook.com"), "bob@outlook.com")

    def test_parse_mailbox_ref_rejects_garbage(self):
        self.assertEqual(service.parse_mailbox_ref("not-a-mailbox"), "")


class TestSecrets(unittest.TestCase):
    def test_hash_and_verify(self):
        stored = store.hash_secret("s3cret", iterations=1000)
        self.assertTrue(stored.startswith("pbkdf2_sha256$1000$"))
        self.assertTrue(store.verify_secret("s3cret", stored))
        self.assertFalse(store.verify_secret("wrong", stored))

    def test_verify_rejects_malformed(self):
        self.assertFalse(store.verify_secret("x", "garbage"))
        self.assertFalse(store.verify_secret("x", ""))

    def test_normalize_scopes(self):
        self.assertEqual(
            store.normalize_scopes(" Mailboxes:Read , otp:read ,, otp:read"),
            ["mailboxes:read", "otp:read"],
        )
        self.assertEqual(store.normalize_scopes(["admin", "admin"]), ["admin"])

    def test_split_token(self):
        self.assertEqual(store.split_token("mbx_sk_abc123_secret", store.KEY_PREFIX),
                         ("abc123", "secret"))
        self.assertEqual(store.split_token("mbx_sk_nosecret", store.KEY_PREFIX), ("", ""))
        self.assertEqual(store.split_token("other", store.KEY_PREFIX), ("", ""))


class TestPrincipalRules(unittest.TestCase):
    def _service(self, scopes, emails=(), all_mailboxes=False, overrides=None):
        return auth.Principal(
            id="p1", kind="service", name="p1", scopes=list(scopes),
            all_mailboxes=all_mailboxes, emails=set(emails),
            scope_overrides=overrides or {},
        )

    def test_wildcard_covers_everything(self):
        p = self._service(["admin", "*"], all_mailboxes=True)
        self.assertTrue(p.has_scope("otp:read"))
        self.assertTrue(p.can_access(MATURE))

    def test_explicit_scopes_only(self):
        p = self._service(["mailboxes:read"], all_mailboxes=True)
        self.assertTrue(p.has_scope("mailboxes:read"))
        self.assertFalse(p.has_scope("fields:sensitive"))

    def test_grant_limits_mailboxes(self):
        p = self._service(["mailboxes:read"], emails={MATURE})
        self.assertTrue(p.can_access(MATURE))
        self.assertFalse(p.can_access(OTHER))

    def test_scope_override_narrows(self):
        p = self._service(
            ["mailboxes:read", "otp:read"], emails={MATURE},
            overrides={MATURE: ["mailboxes:read"]},
        )
        self.assertEqual(p.scopes_for(MATURE), ["mailboxes:read"])

    def test_session_pins_single_mailbox(self):
        p = self._service(["fields:basic"], all_mailboxes=True)
        p.session_email = MATURE
        self.assertTrue(p.can_access(MATURE))
        self.assertFalse(p.can_access(OTHER))

    def test_require_scope_raises(self):
        p = self._service(["mailboxes:read"], all_mailboxes=True)
        with self.assertRaises(MailboxApiError) as ctx:
            auth.require_scope(p, "fields:sensitive")
        self.assertEqual(ctx.exception.status_code, 403)
        self.assertEqual(ctx.exception.code, "scope_required")


class TestFieldSanitizer(unittest.TestCase):
    def setUp(self):
        self.row = account_store.get_account(MATURE)

    def test_basic_view_hides_secrets(self):
        view = service.mailbox_view(self.row, ["mailboxes:read", "fields:basic"])
        self.assertEqual(view["email"], MATURE)
        self.assertIn("batch", view)
        self.assertNotIn("password", view)
        self.assertNotIn("recovery_email", view)

    def test_sensitive_view_adds_secrets(self):
        view = service.mailbox_view(self.row, ["fields:basic", "fields:sensitive"])
        self.assertEqual(view["password"], MATURE_PASSWORD)
        self.assertEqual(view["recovery_email"], "rec@example.com")

    def test_view_never_leaks_refresh_token(self):
        view = service.mailbox_view(self.row, ["fields:basic", "fields:sensitive"])
        self.assertNotIn("refresh_token", view)
        self.assertNotIn("combo", view)

    def test_fields_basic_profile(self):
        out = service.fields_view(self.row, "basic", ["fields:basic"])
        self.assertEqual(
            sorted(out["fields"]), ["batch", "created_at", "email", "incubating", "readable"]
        )

    def test_fields_full_requires_sensitive(self):
        with self.assertRaises(MailboxApiError) as ctx:
            service.fields_view(self.row, "full", ["fields:basic"])
        self.assertEqual(ctx.exception.status_code, 403)

    def test_fields_combo_six_segments(self):
        out = service.fields_view(self.row, "combo", ["fields:basic", "fields:sensitive"])
        self.assertEqual(out["segments"], 6)
        self.assertTrue(out["combo"].startswith(MATURE))

    def test_fields_rejects_unknown_profile(self):
        with self.assertRaises(MailboxApiError) as ctx:
            service.fields_view(self.row, "everything", ["fields:basic", "fields:sensitive"])
        self.assertEqual(ctx.exception.code, "bad_profile")


# ── HTTP ────────────────────────────────────────────────────────────────────


class MailboxApiHttpTest(unittest.TestCase):
    client: TestClient

    @classmethod
    def setUpClass(cls):
        from webapp import server

        cls._ctx = TestClient(server.app)
        cls.client = cls._ctx.__enter__()
        created = cls.client.post(
            "/api/v1/auth/keys",
            headers=ADMIN_HEADERS,
            json={
                "name": "integration",
                "scopes": "mailboxes:read,fields:basic,fields:sensitive,messages:read,otp:read",
            },
        ).json()
        cls.service_key = created["key"]
        cls.service_headers = {"Authorization": f"Bearer {created['key']}"}
        scoped = cls.client.post(
            "/api/v1/auth/keys",
            headers=ADMIN_HEADERS,
            json={"name": "scoped", "scopes": "mailboxes:read", "grants": [MATURE]},
        ).json()
        cls.scoped_headers = {"Authorization": f"Bearer {scoped['key']}"}

    @classmethod
    def tearDownClass(cls):
        cls._ctx.__exit__(None, None, None)

    @property
    def mature_id(self) -> str:
        return service.mailbox_id_for(MATURE)

    @property
    def fresh_id(self) -> str:
        return service.mailbox_id_for(FRESH)

    # 鉴权
    def test_health_requires_token(self):
        r = self.client.get("/api/v1/health")
        self.assertEqual(r.status_code, 401)
        self.assertEqual(r.json()["code"], "unauthorized")

    def test_health_with_service_key(self):
        r = self.client.get("/api/v1/health", headers=self.service_headers)
        self.assertEqual(r.status_code, 200)
        self.assertTrue(r.json()["ok"])
        self.assertGreaterEqual(r.json()["readable_count"], 1)

    def test_me_reports_scopes(self):
        r = self.client.get("/api/v1/auth/me", headers=self.service_headers)
        self.assertEqual(r.status_code, 200)
        self.assertIn("otp:read", r.json()["scopes"])
        self.assertTrue(r.json()["all_mailboxes"])

    def test_create_key_needs_admin(self):
        r = self.client.post(
            "/api/v1/auth/keys", headers=self.service_headers,
            json={"name": "nope", "scopes": "admin"},
        )
        self.assertEqual(r.status_code, 403)
        self.assertEqual(r.json()["required_scope"], "admin")

    def test_revoked_key_stops_working(self):
        created = self.client.post(
            "/api/v1/auth/keys", headers=ADMIN_HEADERS,
            json={"name": "throwaway", "scopes": "mailboxes:read"},
        ).json()
        headers = {"Authorization": f"Bearer {created['key']}"}
        self.assertEqual(self.client.get("/api/v1/health", headers=headers).status_code, 200)
        self.client.delete(f"/api/v1/auth/keys/{created['id']}", headers=ADMIN_HEADERS)
        self.assertEqual(self.client.get("/api/v1/health", headers=headers).status_code, 401)

    # 邮箱列表 / 详情
    def test_list_mailboxes(self):
        r = self.client.get("/api/v1/mailboxes", headers=self.service_headers)
        self.assertEqual(r.status_code, 200)
        body = r.json()
        self.assertGreaterEqual(body["total"], 3)
        emails = {item["email"] for item in body["items"]}
        self.assertIn(MATURE, emails)

    def test_list_mailboxes_pagination_and_filter(self):
        r = self.client.get(
            "/api/v1/mailboxes", headers=self.service_headers, params={"batch": "B2"}
        )
        self.assertEqual([i["email"] for i in r.json()["items"]], [FRESH])
        r2 = self.client.get(
            "/api/v1/mailboxes", headers=self.service_headers, params={"limit": 1, "offset": 0}
        )
        self.assertEqual(len(r2.json()["items"]), 1)

    def test_scoped_key_sees_only_its_grant(self):
        r = self.client.get("/api/v1/mailboxes", headers=self.scoped_headers)
        self.assertEqual([i["email"] for i in r.json()["items"]], [MATURE])
        r2 = self.client.get(f"/api/v1/mailboxes/{self.fresh_id}", headers=self.scoped_headers)
        self.assertEqual(r2.status_code, 403)
        self.assertEqual(r2.json()["code"], "mailbox_forbidden")

    def test_mailbox_detail_and_by_email(self):
        a = self.client.get(f"/api/v1/mailboxes/{self.mature_id}", headers=self.service_headers)
        b = self.client.get(f"/api/v1/mailboxes/by-email/{MATURE}", headers=self.service_headers)
        self.assertEqual(a.status_code, 200)
        self.assertEqual(a.json()["email"], b.json()["email"])

    def test_bad_mailbox_id(self):
        r = self.client.get("/api/v1/mailboxes/mbx_zzz%20zz", headers=self.service_headers)
        self.assertIn(r.status_code, (400, 404))

    def test_fields_basic_over_http(self):
        r = self.client.get(
            f"/api/v1/mailboxes/{self.mature_id}/fields",
            headers=self.service_headers, params={"profile": "basic"},
        )
        self.assertEqual(r.status_code, 200)
        fields = r.json()["fields"]
        self.assertEqual(fields["email"], MATURE)
        self.assertEqual(fields["batch"], "B1")
        self.assertNotIn("password", fields)

    def test_fields_sensitive_blocked_for_scoped_key(self):
        r = self.client.get(
            f"/api/v1/mailboxes/{self.mature_id}/fields",
            headers=self.scoped_headers, params={"profile": "full"},
        )
        self.assertEqual(r.status_code, 403)

    # 用户端会话
    def test_login_and_session_scope(self):
        r = self.client.post(
            "/api/v1/auth/login", json={"email": MATURE, "password": MATURE_PASSWORD}
        )
        self.assertEqual(r.status_code, 200)
        body = r.json()
        self.assertTrue(body["token"].startswith("mbx_sess_"))
        self.assertEqual(body["mailbox_id"], self.mature_id)
        self.assertEqual(body["expires_in"], 2 * 3600)

        headers = {"Authorization": f"Bearer {body['token']}"}
        me = self.client.get("/api/v1/auth/me", headers=headers).json()
        self.assertEqual(me["kind"], "user")
        self.assertEqual(me["session_email"], MATURE)
        self.assertFalse(me["all_mailboxes"])

        own = self.client.get(
            f"/api/v1/mailboxes/{self.mature_id}/fields", headers=headers
        )
        self.assertEqual(own.status_code, 200)
        other = self.client.get(f"/api/v1/mailboxes/{self.fresh_id}/fields", headers=headers)
        self.assertEqual(other.status_code, 403)
        sensitive = self.client.get(
            f"/api/v1/mailboxes/{self.mature_id}/fields",
            headers=headers, params={"profile": "combo"},
        )
        self.assertEqual(sensitive.status_code, 403)

    def test_login_rejects_wrong_password(self):
        r = self.client.post(
            "/api/v1/auth/login", json={"email": MATURE, "password": "nope"}
        )
        self.assertEqual(r.status_code, 401)
        self.assertEqual(r.json()["code"], "invalid_credentials")

    def test_login_rejects_unknown_mailbox(self):
        r = self.client.post(
            "/api/v1/auth/login", json={"email": "ghost@outlook.com", "password": "x"}
        )
        self.assertEqual(r.status_code, 401)

    # 读信
    def test_messages_happy_path(self):
        fake = [{
            "id": "AAMk-1", "subject": "Microsoft account security code",
            "from": "account-security-noreply@accountprotection.microsoft.com",
            "received": "2026-08-28T09:00:00Z", "preview": "Security code: 123456",
            "body": "<p>Security code: 123456</p>",
        }]
        with mock.patch.object(graph_mail, "refresh_token_for",
                               return_value={"access_token": "at", "scope": "Mail.ReadWrite"}), \
             mock.patch.object(graph_mail, "list_messages", return_value=fake) as lm:
            r = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/messages", headers=self.service_headers
            )
        self.assertEqual(r.status_code, 200)
        body = r.json()
        self.assertEqual(body["mode"], "graph")
        self.assertEqual(body["messages"][0]["id"], "AAMk-1")
        self.assertEqual(body["messages"][0]["from"],
                         "account-security-noreply@accountprotection.microsoft.com")
        self.assertEqual(lm.call_args.kwargs["folder"], "inbox")

    def test_messages_filters_by_subject(self):
        fake = [
            {"id": "1", "subject": "hello", "from": "a@b.com", "received": "2026-08-28T09:00:00Z",
             "preview": "", "body": ""},
            {"id": "2", "subject": "your code", "from": "a@b.com",
             "received": "2026-08-28T10:00:00Z", "preview": "", "body": ""},
        ]
        with mock.patch.object(graph_mail, "refresh_token_for",
                               return_value={"access_token": "at", "scope": "Mail.ReadWrite"}), \
             mock.patch.object(graph_mail, "list_messages", return_value=fake):
            r = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/messages",
                headers=self.service_headers, params={"subject_contains": "code"},
            )
        self.assertEqual([m["id"] for m in r.json()["messages"]], ["2"])

    def test_single_message_lookup(self):
        fake = [{"id": "AAMk-9", "subject": "s", "from": "a@b.com",
                 "received": "2026-08-28T09:00:00Z", "preview": "p", "body": "b"}]
        with mock.patch.object(graph_mail, "refresh_token_for",
                               return_value={"access_token": "at", "scope": "Mail.ReadWrite"}), \
             mock.patch.object(graph_mail, "list_messages", return_value=fake):
            hit = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/messages/AAMk-9",
                headers=self.service_headers,
            )
            miss = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/messages/nope",
                headers=self.service_headers,
            )
        self.assertEqual(hit.status_code, 200)
        self.assertEqual(hit.json()["id"], "AAMk-9")
        self.assertEqual(miss.status_code, 404)
        self.assertEqual(miss.json()["code"], "message_not_found")

    def test_otp_returns_code(self):
        fake = [{
            "id": "otp-1", "subject": "Microsoft account security code",
            "from": "account-security-noreply@accountprotection.microsoft.com",
            "received": "2026-08-28T09:00:00Z", "preview": "Security code: 778899",
            "body": "<p>Security code: 778899</p>",
        }]
        with mock.patch.object(graph_mail, "refresh_token_for",
                               return_value={"access_token": "at", "scope": "Mail.ReadWrite"}), \
             mock.patch.object(graph_mail, "list_messages", return_value=fake):
            r = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/otp",
                headers=self.service_headers, params={"wait_seconds": 0},
            )
        self.assertEqual(r.status_code, 200)
        self.assertTrue(r.json()["found"])
        self.assertEqual(r.json()["code"], "778899")

    def test_otp_reports_miss_without_blocking(self):
        with mock.patch.object(graph_mail, "refresh_token_for",
                               return_value={"access_token": "at", "scope": "Mail.ReadWrite"}), \
             mock.patch.object(graph_mail, "list_messages", return_value=[]):
            r = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/otp",
                headers=self.service_headers, params={"wait_seconds": 0},
            )
        self.assertEqual(r.status_code, 200)
        self.assertFalse(r.json()["found"])
        self.assertEqual(r.json()["code"], "")

    def test_otp_wait_seconds_capped(self):
        r = self.client.get(
            f"/api/v1/mailboxes/{self.mature_id}/otp",
            headers=self.service_headers, params={"wait_seconds": 600},
        )
        self.assertEqual(r.status_code, 422)

    # 业务规则
    def test_incubating_mailbox_is_locked(self):
        r = self.client.get(
            f"/api/v1/mailboxes/{self.fresh_id}/messages", headers=self.service_headers
        )
        self.assertEqual(r.status_code, 423)
        body = r.json()
        self.assertEqual(body["code"], "incubating")
        self.assertTrue(body["until"])

    def test_incubating_bypass_for_admin(self):
        with mock.patch.object(graph_mail, "refresh_token_for",
                               return_value={"access_token": "at", "scope": "Mail.ReadWrite"}), \
             mock.patch.object(graph_mail, "list_messages", return_value=[]):
            r = self.client.get(
                f"/api/v1/mailboxes/{self.fresh_id}/messages", headers=ADMIN_HEADERS
            )
        self.assertEqual(r.status_code, 200)

    def test_mailbox_without_token(self):
        r = self.client.get(
            f"/api/v1/mailboxes/{service.mailbox_id_for(OTHER)}/messages",
            headers=self.service_headers,
        )
        self.assertEqual(r.status_code, 404)
        self.assertEqual(r.json()["code"], "no_token")

    def test_dead_token_is_explicit(self):
        with mock.patch.object(graph_mail, "refresh_token_for",
                               return_value={"error": "invalid_grant",
                                             "error_description": "token expired"}):
            r = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/messages", headers=self.service_headers
            )
        self.assertEqual(r.status_code, 502)
        self.assertEqual(r.json()["code"], "token_dead")

    def test_transient_upstream_is_not_dead(self):
        with mock.patch.object(graph_mail, "refresh_token_for",
                               return_value={"error": "transient_network", "transient": True}):
            r = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/messages", headers=self.service_headers
            )
        self.assertEqual(r.status_code, 503)
        self.assertEqual(r.json()["code"], "upstream_unavailable")

    def test_rotated_refresh_token_is_written_back(self):
        with mock.patch.object(
            graph_mail, "refresh_token_for",
            return_value={"access_token": "at", "refresh_token": "rt-rotated",
                          "scope": "Mail.ReadWrite"},
        ), mock.patch.object(graph_mail, "list_messages", return_value=[]):
            r = self.client.get(
                f"/api/v1/mailboxes/{self.mature_id}/messages", headers=self.service_headers
            )
        self.assertEqual(r.status_code, 200)
        row = account_store.get_account(MATURE)
        self.assertEqual(row["refresh_token"], "rt-rotated")
        self.assertEqual(row["combo"].split("----")[3], "rt-rotated")
        self.assertEqual(row["combo_recovery"].split("----")[3], "rt-rotated")
        # 复原，避免影响其他用例
        account_store.patch_account(MATURE, {
            "refresh_token": "rt-mature",
            "combo": f"{MATURE}----{MATURE_PASSWORD}----cid----rt-mature",
            "combo_recovery": f"{MATURE}----{MATURE_PASSWORD}----cid----rt-mature"
                              "----rec@example.com----recpwd",
        })

    # 审计
    def test_audit_records_requests_without_tokens(self):
        self.client.get("/api/v1/health", headers=self.service_headers)
        rows = store.recent_audit(50)
        self.assertTrue(rows)
        paths = {r["path"] for r in rows}
        self.assertIn("/api/v1/health", paths)
        blob = " ".join(str(r) for r in rows)
        self.assertNotIn(self.service_key, blob)
        self.assertNotIn("test-admin-key", blob)


if __name__ == "__main__":
    unittest.main()
