"""`/api/accounts/import` 与 `/api/accounts/import/qoderji` 的 HTTP 端到端行为。"""
from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import time
import unittest
from pathlib import Path

_TMP_DIR = tempfile.mkdtemp(prefix="import-endpoints-test-")
_PREV_ENV = {
    k: os.environ.get(k)
    for k in (
        "OUTLOOK_DB_PATH",
        "OUTLOOK_MAILBOX_API_ENABLED",
        "QODERJI_EMAIL_DB",
        "OUTLOOK_CONSOLE_PASSWORD",
    )
}
os.environ["OUTLOOK_DB_PATH"] = os.path.join(_TMP_DIR, "outlook.db")
os.environ["OUTLOOK_MAILBOX_API_ENABLED"] = "1"
os.environ.pop("QODERJI_EMAIL_DB", None)
# 运维台登录门只在口令非空时生效；这里测的是账号导入接口本身，不测登录门
# （login gate 已有 test_console_auth.py 专门覆盖），显式清空以免开发机
# shell 里残留的 OUTLOOK_CONSOLE_PASSWORD 把这些请求挡在 401。
os.environ["OUTLOOK_CONSOLE_PASSWORD"] = ""

from fastapi.testclient import TestClient  # noqa: E402

from outlook_api_reg import account_store, database as db  # noqa: E402


def setUpModule() -> None:
    db._initialized = False
    db.ensure_initialized()
    conn = db.connect()
    try:
        db._app_set(conn, "legacy_files_migrated", "1")
        conn.commit()
    finally:
        conn.close()


def tearDownModule() -> None:
    for key, value in _PREV_ENV.items():
        if value is None:
            os.environ.pop(key, None)
        else:
            os.environ[key] = value
    db._initialized = False
    shutil.rmtree(_TMP_DIR, ignore_errors=True)


def _clear_accounts() -> None:
    conn = db.connect()
    try:
        conn.execute("DELETE FROM accounts")
        conn.execute("DELETE FROM account_meta")
        conn.commit()
    finally:
        conn.close()


def _make_qoderji_db(path: Path, rows: list[dict]) -> None:
    conn = sqlite3.connect(str(path))
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS email_inventory (
                email TEXT PRIMARY KEY,
                raw TEXT NOT NULL DEFAULT '',
                batch_id TEXT NOT NULL DEFAULT '',
                source TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'untried',
                added_at REAL NOT NULL
            );
            """
        )
        now = time.time()
        for r in rows:
            conn.execute(
                "INSERT INTO email_inventory(email, raw, status, added_at) VALUES (?,?,?,?)",
                (r["email"], r["raw"], r.get("status", "untried"), r.get("added_at", now)),
            )
        conn.commit()
    finally:
        conn.close()


class ImportEndpointsTest(unittest.TestCase):
    client: TestClient

    @classmethod
    def setUpClass(cls) -> None:
        from webapp import server

        cls._ctx = TestClient(server.app)
        cls.client = cls._ctx.__enter__()

    @classmethod
    def tearDownClass(cls) -> None:
        cls._ctx.__exit__(None, None, None)

    def setUp(self) -> None:
        _clear_accounts()

    def test_paste_import_backward_compatible_response_shape(self):
        text = "p1@outlook.com----pwd----cid----rt1\np2@outlook.com----pwd----cid----rt2----lcid----lrt\n"
        r = self.client.post("/api/accounts/import", json={"text": text})
        self.assertEqual(r.status_code, 200)
        body = r.json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["imported"], 2)
        self.assertEqual(body["duplicate"], 0)
        self.assertEqual(body["invalid"], 0)
        self.assertEqual(body["six_seg"], 1)

    def test_paste_import_with_source_and_skip_incubation(self):
        text = "ext1@outlook.com----pwd----cid----rt1\n"
        r = self.client.post(
            "/api/accounts/import",
            json={"text": text, "source": "manual-merge", "batch_label": "手工合并", "skip_incubation": True},
        )
        self.assertEqual(r.status_code, 200)
        body = r.json()
        self.assertEqual(body["imported"], 1)
        self.assertEqual(body["source"], "manual-merge")
        self.assertEqual(body["batch_label"], "手工合并")
        self.assertTrue(body["skip_incubation"])

        detail = self.client.get("/api/accounts").json()
        row = next(a for a in detail["accounts"] if a["email"] == "ext1@outlook.com")
        self.assertFalse(row["incubating"])
        self.assertIn("src:manual-merge", row["tags"])

    def test_paste_import_dedupe_across_calls(self):
        text = "dup@outlook.com----pwd----cid----rt1\n"
        r1 = self.client.post("/api/accounts/import", json={"text": text})
        r2 = self.client.post("/api/accounts/import", json={"text": text})
        self.assertEqual(r1.json()["imported"], 1)
        self.assertEqual(r2.json()["duplicate"], 1)
        self.assertEqual(r2.json()["imported"], 0)

    def test_qoderji_import_happy_path(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": "q1@outlook.com", "raw": "q1@outlook.com----x----cid----rt1", "status": "consumed"},
                {"email": "qdead@outlook.com", "raw": "qdead@outlook.com----x----cid----rt2", "status": "dead"},
            ])
            r = self.client.post("/api/accounts/import/qoderji", json={"db_path": str(dbp)})
            self.assertEqual(r.status_code, 200)
            body = r.json()
            self.assertTrue(body["ok"])
            self.assertEqual(body["imported"], 1)
            self.assertEqual(body["source"], "qoderji")

            detail = self.client.get("/api/accounts").json()
            emails = {a["email"] for a in detail["accounts"]}
            self.assertIn("q1@outlook.com", emails)
            self.assertNotIn("qdead@outlook.com", emails)
            row = next(a for a in detail["accounts"] if a["email"] == "q1@outlook.com")
            self.assertFalse(row["incubating"])  # 默认 skip_incubation=True

    def test_qoderji_import_404_when_no_db_found(self):
        with tempfile.TemporaryDirectory() as d:
            r = self.client.post(
                "/api/accounts/import/qoderji",
                json={"db_path": str(Path(d) / "does-not-exist.db")},
            )
            self.assertEqual(r.status_code, 404)

    def test_qoderji_import_status_filter_and_dry_run(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": "untried1@outlook.com", "raw": "untried1@outlook.com----x----cid----rt1", "status": "untried"},
                {"email": "consumed1@outlook.com", "raw": "consumed1@outlook.com----x----cid----rt2", "status": "consumed"},
            ])
            r = self.client.post(
                "/api/accounts/import/qoderji",
                json={"db_path": str(dbp), "statuses": ["consumed"], "dry_run": True},
            )
            body = r.json()
            self.assertEqual(body["imported"], 1)
            detail = self.client.get("/api/accounts").json()
            emails = {a["email"] for a in detail["accounts"]}
            self.assertNotIn("consumed1@outlook.com", emails)  # dry_run 不写库


if __name__ == "__main__":
    unittest.main()
