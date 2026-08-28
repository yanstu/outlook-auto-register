"""外部账号合并（external_import）：combo 解析、去重导入、qoderji 只读拉取。"""
from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

_TMP_DIR = tempfile.mkdtemp(prefix="external-import-test-")
_PREV_ENV = {
    k: os.environ.get(k)
    for k in ("OUTLOOK_DB_PATH", "OUTLOOK_INCUBATION_HOURS", "QODERJI_EMAIL_DB")
}
os.environ["OUTLOOK_DB_PATH"] = os.path.join(_TMP_DIR, "outlook.db")
os.environ["OUTLOOK_INCUBATION_HOURS"] = "48"
os.environ.pop("QODERJI_EMAIL_DB", None)

from outlook_api_reg import account_store, database as db, external_import, lifecycle  # noqa: E402


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
    """按 qoderji email_pool.py 的真实建表语句起一份最小可用的 email_inventory。"""
    conn = sqlite3.connect(str(path))
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS email_inventory (
                email        TEXT PRIMARY KEY,
                raw          TEXT NOT NULL DEFAULT '',
                batch_id     TEXT NOT NULL DEFAULT '',
                source       TEXT NOT NULL DEFAULT '',
                status       TEXT NOT NULL DEFAULT 'untried',
                leased_by    TEXT NOT NULL DEFAULT '',
                leased_name  TEXT NOT NULL DEFAULT '',
                lease_job    TEXT NOT NULL DEFAULT '',
                lease_expires_at REAL,
                dead_reason  TEXT,
                added_at     REAL NOT NULL,
                leased_at    REAL,
                consumed_at  REAL
            );
            """
        )
        now = time.time()
        for r in rows:
            conn.execute(
                "INSERT INTO email_inventory(email, raw, batch_id, source, status, added_at) "
                "VALUES (?,?,?,?,?,?)",
                (
                    r["email"],
                    r.get("raw", ""),
                    r.get("batch_id", ""),
                    r.get("source", ""),
                    r.get("status", "untried"),
                    r.get("added_at", now),
                ),
            )
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# 纯函数：combo 解析
# ---------------------------------------------------------------------------


class TestParseComboLine(unittest.TestCase):
    def test_four_segment(self):
        c = external_import.parse_combo_line("a@outlook.com----pwd----cid1----rt1")
        self.assertIsNotNone(c)
        self.assertEqual(c.email, "a@outlook.com")
        self.assertEqual(c.password, "pwd")
        self.assertEqual(c.client_id, "cid1")
        self.assertEqual(c.refresh_token, "rt1")
        self.assertEqual(c.combo4(), "a@outlook.com----pwd----cid1----rt1")

    def test_six_segment_dual(self):
        c = external_import.parse_combo_line(
            "a@outlook.com----pwd----cid1----rt1----logincid----loginrt"
        )
        self.assertIsNotNone(c)
        self.assertEqual(c.login_client_id, "logincid")
        self.assertEqual(c.login_refresh_token, "loginrt")
        self.assertEqual(
            c.combo6_dual(), "a@outlook.com----pwd----cid1----rt1----logincid----loginrt"
        )

    def test_qoderji_placeholder_password(self):
        # qoderji 的 raw 常见形态：password 段是占位符 "x"，真正靠 refresh_token。
        c = external_import.parse_combo_line("bob@outlook.com----x----9e5f94bc----M.C5_abc.def")
        self.assertIsNotNone(c)
        self.assertEqual(c.password, "x")
        self.assertEqual(c.refresh_token, "M.C5_abc.def")

    def test_bare_email_no_token(self):
        c = external_import.parse_combo_line("bare@outlook.com")
        self.assertIsNotNone(c)
        self.assertEqual(c.email, "bare@outlook.com")
        self.assertEqual(c.refresh_token, "")

    def test_invalid_no_at(self):
        self.assertIsNone(external_import.parse_combo_line("not-an-email----pwd----cid----rt"))

    def test_blank_and_comment_ignored(self):
        self.assertIsNone(external_import.parse_combo_line(""))
        self.assertIsNone(external_import.parse_combo_line("   "))
        self.assertIsNone(external_import.parse_combo_line("# comment----x----y----z"))

    def test_pipe_separated_fallback(self):
        c = external_import.parse_combo_line("weird@outlook.com|somepwd")
        self.assertIsNotNone(c)
        self.assertEqual(c.email, "weird@outlook.com")


class TestParseComboText(unittest.TestCase):
    def test_mixed_valid_invalid(self):
        text = "\n".join([
            "a@outlook.com----pwd----cid----rt1",
            "not-an-email",
            "",
            "# comment",
            "b@outlook.com----pwd----cid----rt2",
        ])
        combos, stats = external_import.parse_combo_text(text)
        self.assertEqual(len(combos), 2)
        self.assertEqual(stats["invalid"], 1)


# ---------------------------------------------------------------------------
# import_combos / import_text：去重 + 孵化期回填
# ---------------------------------------------------------------------------


class TestImportCombos(unittest.TestCase):
    def setUp(self) -> None:
        _clear_accounts()

    def test_import_fresh_combos(self):
        combos = [
            external_import.ExternalCombo(email="new1@outlook.com", password="p", client_id="c", refresh_token="rt1"),
            external_import.ExternalCombo(email="new2@outlook.com", password="p", client_id="c", refresh_token="rt2"),
        ]
        result = external_import.import_combos(combos, source="unittest", batch_label="批次A")
        self.assertEqual(result["imported"], 2)
        self.assertEqual(result["duplicate"], 0)
        self.assertEqual(result["invalid"], 0)
        acc = account_store.get_account("new1@outlook.com")
        self.assertIsNotNone(acc)
        self.assertEqual(acc["refresh_token"], "rt1")
        self.assertEqual(acc["batch_label"], "批次A")
        self.assertIn("src:unittest", acc["tags"])

    def test_dedupe_against_existing_account(self):
        conn = db.connect()
        try:
            account_store.upsert_account_dict(conn, {
                "email": "dup@outlook.com",
                "password": "old",
                "client_id": "old-cid",
                "refresh_token": "old-rt",
                "combo": "dup@outlook.com----old----old-cid----old-rt",
                "created_at": "2020-01-01T00:00:00",
                "updated_at": "2020-01-01T00:00:00",
            })
            conn.commit()
        finally:
            conn.close()
        combos = [
            external_import.ExternalCombo(email="dup@outlook.com", password="new", client_id="new-cid", refresh_token="new-rt"),
            external_import.ExternalCombo(email="fresh@outlook.com", password="p", client_id="c", refresh_token="rt"),
        ]
        result = external_import.import_combos(combos, source="unittest")
        self.assertEqual(result["duplicate"], 1)
        self.assertEqual(result["imported"], 1)
        # 已存在账号的字段不应被外部导入覆盖
        acc = account_store.get_account("dup@outlook.com")
        self.assertEqual(acc["refresh_token"], "old-rt")

    def test_dedupe_within_same_batch(self):
        combos = [
            external_import.ExternalCombo(email="A@outlook.com", refresh_token="rt1"),
            external_import.ExternalCombo(email="a@outlook.com", refresh_token="rt2"),
        ]
        result = external_import.import_combos(combos, source="unittest")
        self.assertEqual(result["imported"], 1)
        self.assertEqual(result["duplicate"], 1)

    def test_missing_refresh_token_is_invalid(self):
        combos = [external_import.ExternalCombo(email="notoken@outlook.com", password="p")]
        result = external_import.import_combos(combos, source="unittest")
        self.assertEqual(result["imported"], 0)
        self.assertEqual(result["invalid"], 1)
        self.assertIsNone(account_store.get_account("notoken@outlook.com"))

    def test_skip_incubation_backdates_created_at(self):
        combos = [external_import.ExternalCombo(email="longlived@outlook.com", refresh_token="rt")]
        external_import.import_combos(combos, source="unittest", skip_incubation=True)
        acc = account_store.get_account("longlived@outlook.com")
        self.assertFalse(acc["incubating"])
        self.assertNotIn(lifecycle.INCUBATING_TAG, acc["tags"])

    def test_without_skip_incubation_stays_incubating(self):
        combos = [external_import.ExternalCombo(email="freshimport@outlook.com", refresh_token="rt")]
        external_import.import_combos(combos, source="unittest", skip_incubation=False)
        acc = account_store.get_account("freshimport@outlook.com")
        self.assertTrue(acc["incubating"])

    def test_dry_run_does_not_write(self):
        combos = [external_import.ExternalCombo(email="dryrun@outlook.com", refresh_token="rt")]
        result = external_import.import_combos(combos, source="unittest", dry_run=True)
        self.assertEqual(result["imported"], 1)
        self.assertIsNone(account_store.get_account("dryrun@outlook.com"))

    def test_import_text_matches_paste_semantics(self):
        text = "p1@outlook.com----pwd----cid----rt1----logincid----loginrt\nbadline\n"
        result = external_import.import_text(text, source="paste")
        self.assertEqual(result["imported"], 1)
        self.assertEqual(result["six_seg"], 1)
        self.assertEqual(result["invalid"], 1)


# ---------------------------------------------------------------------------
# qoderji：候选路径探测 + 只读拉取
# ---------------------------------------------------------------------------


class TestQoderjiDbCandidates(unittest.TestCase):
    def test_explicit_path(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "one.db"
            p.write_bytes(b"")
            found = external_import.qoderji_db_candidates(str(p))
            self.assertEqual(len(found), 1)
            self.assertEqual(found[0].resolve(), p.resolve())

    def test_explicit_comma_separated(self):
        with tempfile.TemporaryDirectory() as d:
            p1 = Path(d) / "a.db"
            p2 = Path(d) / "b.db"
            p1.write_bytes(b"")
            p2.write_bytes(b"")
            found = external_import.qoderji_db_candidates(f"{p1},{p2}")
            self.assertEqual({f.resolve() for f in found}, {p1.resolve(), p2.resolve()})

    def test_missing_returns_empty(self):
        with tempfile.TemporaryDirectory() as d:
            found = external_import.qoderji_db_candidates(str(Path(d) / "nope.db"))
            self.assertEqual(found, [])

    def test_default_glob(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "cards.db"
            p.write_bytes(b"")
            with mock.patch.object(external_import, "DEFAULT_QODERJI_DB_GLOBS", (str(Path(d) / "*.db"),)):
                found = external_import.qoderji_db_candidates(None)
            self.assertEqual(len(found), 1)


class TestFetchQoderjiCombos(unittest.TestCase):
    def test_excludes_dead_by_default(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": "u1@outlook.com", "raw": "u1@outlook.com----x----cid1----rt1", "status": "untried"},
                {"email": "c1@outlook.com", "raw": "c1@outlook.com----x----cid2----rt2", "status": "consumed"},
                {"email": "l1@outlook.com", "raw": "l1@outlook.com----x----cid3----rt3", "status": "leased"},
                {"email": "d1@outlook.com", "raw": "d1@outlook.com----x----cid4----rt4", "status": "dead"},
            ])
            combos, stats = external_import.fetch_qoderji_combos(dbp)
            emails = {c.email for c in combos}
            self.assertEqual(emails, {"u1@outlook.com", "c1@outlook.com", "l1@outlook.com"})
            self.assertEqual(stats["scanned"], 3)
            self.assertNotIn("dead", stats["by_status"])

    def test_status_filter(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": "u1@outlook.com", "raw": "u1@outlook.com----x----cid----rt1", "status": "untried"},
                {"email": "c1@outlook.com", "raw": "c1@outlook.com----x----cid----rt2", "status": "consumed"},
            ])
            combos, _stats = external_import.fetch_qoderji_combos(dbp, statuses=("consumed",))
            self.assertEqual({c.email for c in combos}, {"c1@outlook.com"})

    def test_malformed_raw_counts_invalid(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": "bare@outlook.com", "raw": "bare@outlook.com", "status": "untried"},
            ])
            combos, stats = external_import.fetch_qoderji_combos(dbp)
            self.assertEqual(combos, [])
            self.assertEqual(stats["invalid"], 1)

    def test_extra_metadata_carried(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {
                    "email": "meta@outlook.com",
                    "raw": "meta@outlook.com----x----cid----rt",
                    "status": "consumed",
                    "batch_id": "batch-123",
                    "source": "order.txt",
                },
            ])
            combos, _stats = external_import.fetch_qoderji_combos(dbp)
            self.assertEqual(len(combos), 1)
            self.assertEqual(combos[0].extra["qoderji_status"], "consumed")
            self.assertEqual(combos[0].extra["qoderji_batch_id"], "batch-123")
            self.assertEqual(combos[0].extra["qoderji_source"], "order.txt")

    def test_missing_table_reports_error(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "empty.db"
            sqlite3.connect(str(dbp)).close()
            combos, stats = external_import.fetch_qoderji_combos(dbp)
            self.assertEqual(combos, [])
            self.assertIn("error", stats)

    def test_survives_locked_wal_directory(self):
        """复现线上真实故障：qoderji 的库是 WAL 模式，数据目录只有 root 能写。

        普通 ``mode=ro`` 连接在 WAL 库上仍可能要创建/更新 ``-shm`` 才能读，没有目录
        写权限时会炸 ``attempt to write a readonly database``；这里验证
        `_open_readonly` 的 ``immutable=1`` 兜底能在这种目录下依然读出数据。
        """
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "wal.db"
            conn = sqlite3.connect(str(dbp))
            conn.execute("PRAGMA journal_mode=WAL")
            conn.executescript(
                """
                CREATE TABLE email_inventory (
                    email TEXT PRIMARY KEY, raw TEXT NOT NULL DEFAULT '',
                    batch_id TEXT NOT NULL DEFAULT '', source TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT 'untried', added_at REAL NOT NULL
                );
                """
            )
            conn.execute(
                "INSERT INTO email_inventory(email, raw, status, added_at) VALUES (?,?,?,?)",
                ("wal1@outlook.com", "wal1@outlook.com----x----cid----rt", "untried", time.time()),
            )
            conn.commit()
            conn.close()  # 干净关闭：WAL 自动 checkpoint，-wal/-shm 通常被清掉（复现线上现场）
            for suffix in ("-wal", "-shm"):
                p = Path(str(dbp) + suffix)
                if p.exists():
                    p.unlink()
            os.chmod(d, 0o555)  # 模拟目录属主是 root、我们的进程用户没有写权限
            try:
                combos, stats = external_import.fetch_qoderji_combos(dbp)
            finally:
                os.chmod(d, 0o755)
            self.assertNotIn("error", stats)
            self.assertEqual({c.email for c in combos}, {"wal1@outlook.com"})

    def test_readonly_does_not_lock_writer(self):
        """确认只读连接不会阻塞 qoderji 自己继续写这份库（同机常驻服务的现实约束）。"""
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": "u1@outlook.com", "raw": "u1@outlook.com----x----cid----rt1", "status": "untried"},
            ])
            external_import.fetch_qoderji_combos(dbp)
            # 只读连接必须已经关闭/不持有写锁：这里还能正常写入才算过关
            conn = sqlite3.connect(str(dbp))
            conn.execute(
                "INSERT INTO email_inventory(email, raw, status, added_at) VALUES (?,?,?,?)",
                ("u2@outlook.com", "u2@outlook.com----x----cid----rt2", "untried", time.time()),
            )
            conn.commit()
            conn.close()


class TestImportFromQoderji(unittest.TestCase):
    def setUp(self) -> None:
        _clear_accounts()

    def test_full_pipeline_skips_incubation_by_default(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": "q1@outlook.com", "raw": "q1@outlook.com----x----cid----rt1", "status": "consumed"},
                {"email": "q2@outlook.com", "raw": "q2@outlook.com----x----cid----rt2", "status": "untried"},
                {"email": "qdead@outlook.com", "raw": "qdead@outlook.com----x----cid----rt3", "status": "dead"},
            ])
            result = external_import.import_from_qoderji(db_path=str(dbp))
            self.assertTrue(result["ok"])
            self.assertEqual(result["imported"], 2)
            self.assertEqual(result["source"], "qoderji")
            acc = account_store.get_account("q1@outlook.com")
            self.assertIsNotNone(acc)
            self.assertFalse(acc["incubating"])
            self.assertIsNone(account_store.get_account("qdead@outlook.com"))

    def test_no_db_found_returns_error(self):
        with tempfile.TemporaryDirectory() as d:
            result = external_import.import_from_qoderji(db_path=str(Path(d) / "nope.db"))
            self.assertFalse(result["ok"])
            self.assertIn("error", result)

    def test_dry_run_reports_without_writing(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": "qdry@outlook.com", "raw": "qdry@outlook.com----x----cid----rt", "status": "untried"},
            ])
            result = external_import.import_from_qoderji(db_path=str(dbp), dry_run=True)
            self.assertEqual(result["imported"], 1)
            self.assertIsNone(account_store.get_account("qdry@outlook.com"))

    def test_limit_caps_fetched(self):
        with tempfile.TemporaryDirectory() as d:
            dbp = Path(d) / "cards.db"
            _make_qoderji_db(dbp, [
                {"email": f"lim{i}@outlook.com", "raw": f"lim{i}@outlook.com----x----cid----rt{i}", "status": "untried"}
                for i in range(5)
            ])
            result = external_import.import_from_qoderji(db_path=str(dbp), limit=2)
            self.assertEqual(result["fetched"], 2)
            self.assertEqual(result["imported"], 2)


if __name__ == "__main__":
    unittest.main()
