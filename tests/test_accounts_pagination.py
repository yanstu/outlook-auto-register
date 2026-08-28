"""账号池分页 / 瘦身字段 / 概览统计的正确性与体量回归测试。

背景：`/api/accounts` 曾一次性把全表（近万条，约 22MB）搬给浏览器，导致页面加载
失败、界面显示空号池。修复后：服务端分页（limit/offset + q/batch/view 过滤）+
列表行瘦身（不带完整 refresh_token/combo，只给 has_token/token_tail），完整明文
只在 `/api/accounts/{email}` 详情接口给。

本文件覆盖：
- account_store 层：count_accounts / list_accounts_page 的过滤口径与旧版
  list_accounts()（Python 侧全量过滤）逐条比对，分页遍历不重不漏。
- overview_stats()：总数/有令牌/可用/失活/未测/批次分布。
- HTTP 层：GET /api/accounts 分页响应形状 + 列表行确实不带 refresh_token/combo；
  GET /api/accounts/{email} 单号详情带完整字段；GET /api/accounts/stats。
- 体量回归：模拟近千条大 token 账号，默认 limit=50 的响应体积仍是 KB 级，不是
  之前那种随总量线性膨胀的 MB 级。
"""
from __future__ import annotations

import json
import os
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta

_TMP_DIR = tempfile.mkdtemp(prefix="accounts-pagination-test-")
# 注意：OUTLOOK_CONSOLE_PASSWORD 故意不放进 _PREV_ENV 去「恢复原值」——这个进程级
# 环境变量在整个 pytest 会话里跨测试文件共享，若在 tearDownModule 里把它恢复成
# 收集阶段之前 shell 里的真实口令，会在本文件测试跑完、后面别的测试文件（如
# test_import_endpoints.py）开始跑之前，把 401 门重新关上，污染它们。跟
# test_import_endpoints.py 的约定一致：清空后就一直留空，不恢复。
_PREV_ENV = {k: os.environ.get(k) for k in ("OUTLOOK_DB_PATH", "OUTLOOK_MAILBOX_API_ENABLED")}
os.environ["OUTLOOK_DB_PATH"] = os.path.join(_TMP_DIR, "outlook.db")
os.environ["OUTLOOK_CONSOLE_PASSWORD"] = ""
os.environ["OUTLOOK_MAILBOX_API_ENABLED"] = "0"

from fastapi.testclient import TestClient  # noqa: E402

from outlook_api_reg import account_store  # noqa: E402
from outlook_api_reg import database as db  # noqa: E402

# 孵化期默认 48h 会让刚插入的账号一直显示为「孵化中」，与用于对拍的旧版
# list_accounts() 过滤口径无关，但拉长 created_at 更贴近真实号池，也让
# incubating 计算走一次真实分支。
_OLD_CREATED = (datetime.now() - timedelta(days=30)).isoformat(timespec="seconds")


def setUpModule() -> None:
    db._initialized = False
    db.ensure_initialized()


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
        conn.execute("DELETE FROM rescue_events")
        conn.commit()
    finally:
        conn.close()


def _seed_account(
    email: str,
    *,
    has_token: bool = True,
    token_len: int = 1200,
    recovery: bool = False,
    batch_no: "int | None" = None,
    batch_label: str = "",
    note: str = "",
    tags: "list[str] | None" = None,
    verify: "dict | None" = None,
    created_at: str = _OLD_CREATED,
) -> None:
    """插一条合成账号，绕开真正的注册流程，只为覆盖分页/过滤逻辑。"""
    refresh_token = ("RT." + "x" * token_len) if has_token else ""
    row = {
        "email": email,
        "password": f"pwd-{email}",
        "client_id": "9e5f94bc-e8a4-4e73-b8be-63364c29d753",
        "refresh_token": refresh_token,
        "combo": "----".join([email, f"pwd-{email}", "cid", refresh_token]),
        "recovery_email": f"rec-{email}" if recovery else "",
        "recovery_password": "cf_domain" if recovery else "",
        "batch_id": f"job-{batch_no}" if batch_no else "",
        "batch_no": batch_no,
        "batch_label": batch_label,
        "created_at": created_at,
        "updated_at": created_at,
    }
    conn = db.connect()
    try:
        account_store.upsert_account_dict(conn, row)
        if note or tags or verify is not None:
            account_store.upsert_meta_dict(
                conn,
                email,
                {"note": note, "tags": tags or [], "verify": verify, "updated_at": created_at},
            )
        conn.commit()
    finally:
        conn.close()


class AccountStoreFilterTest(unittest.TestCase):
    """count_accounts / list_accounts_page 的 SQL 过滤要跟旧版 Python 过滤同口径。"""

    def setUp(self) -> None:
        _clear_accounts()
        for i in range(40):
            has_token = i % 3 != 0
            verify = None
            if has_token and i % 2 == 0:
                verify = {"ok": i % 5 != 0, "graph": i % 5 != 0}
            _seed_account(
                f"user{i:03d}@outlook.com",
                has_token=has_token,
                recovery=(i % 4 == 0),
                batch_no=(i % 5) + 1,
                batch_label=f"B{(i % 5) + 1}",
                note=("vip" if i == 7 else ""),
                tags=(["hot"] if i % 10 == 0 else []),
                verify=verify,
            )
        # 一条无批次账号，专门盖「无批次」分支
        _seed_account("nobatch@outlook.com", has_token=True)

    def _expected(self, *, view: str = "all", batch: str = "all", q: str = "") -> int:
        rows = account_store.list_accounts()
        out = 0
        for r in rows:
            v = r.get("verify")
            is_usable = bool(v and (v.get("ok") or v.get("graph")))
            if view == "without" and r.get("has_token"):
                continue
            if view == "dual" and not r.get("login_token"):
                continue
            if view == "untested" and v is not None:
                continue
            if view == "usable" and not is_usable:
                continue
            if view == "dead" and not (v and not v.get("ok") and not v.get("graph")):
                continue
            label = r.get("batch_label") or (f"B{r['batch_no']}" if r.get("batch_no") else "")
            if batch == "none" and label:
                continue
            if batch not in ("all", "none") and label != batch:
                continue
            if q:
                hay = " ".join(
                    [r["email"], r.get("recovery_email") or "", r.get("note") or "", " ".join(r.get("tags") or []), label]
                ).lower()
                if q.lower() not in hay:
                    continue
            out += 1
        return out

    def test_view_filters_match_legacy_python_filter(self) -> None:
        for view in ("all", "without", "dual", "untested", "usable", "dead"):
            with self.subTest(view=view):
                self.assertEqual(
                    account_store.count_accounts({"view": view}),
                    self._expected(view=view),
                )

    def test_batch_filters_match_legacy_python_filter(self) -> None:
        for batch in ("all", "none", "B1", "B3", "B5"):
            with self.subTest(batch=batch):
                self.assertEqual(
                    account_store.count_accounts({"batch": batch}),
                    self._expected(batch=batch),
                )

    def test_search_filters_match_legacy_python_filter(self) -> None:
        for q in ("user001", "vip", "hot", "rec-user000", "zz-no-match"):
            with self.subTest(q=q):
                self.assertEqual(
                    account_store.count_accounts({"q": q}),
                    self._expected(q=q),
                )

    def test_pagination_covers_every_row_without_duplicates(self) -> None:
        total = account_store.count_accounts({})
        seen: set[str] = set()
        offset = 0
        page_size = 7  # 故意用不能整除总数的页大小，逼出边界情况
        while offset < total:
            page = account_store.list_accounts_page(limit=page_size, offset=offset)
            self.assertLessEqual(len(page), page_size)
            for row in page:
                self.assertNotIn(row["email"], seen, "分页出现重复邮箱")
                seen.add(row["email"])
            offset += page_size
        self.assertEqual(len(seen), total)

    def test_sort_order_matches_legacy_usable_first_then_created_desc(self) -> None:
        rows = account_store.list_accounts()

        def is_usable(r: dict) -> bool:
            v = r.get("verify")
            return bool(v and (v.get("ok") or v.get("graph")))

        expected = sorted(rows, key=lambda r: (1 if is_usable(r) else 0, r.get("created_at") or ""), reverse=True)
        expected_emails = [r["email"] for r in expected]

        got_emails: list[str] = []
        total = account_store.count_accounts({})
        offset = 0
        while offset < total:
            page = account_store.list_accounts_page(limit=10, offset=offset)
            got_emails.extend(r["email"] for r in page)
            offset += 10
        self.assertEqual(got_emails, expected_emails)

    def test_list_rows_are_slimmed_no_full_secrets(self) -> None:
        page = account_store.list_accounts_page(limit=100, offset=0)
        for row in page:
            self.assertNotIn("refresh_token", row)
            self.assertNotIn("combo", row)
            self.assertNotIn("combo_dual", row)
            self.assertNotIn("login_refresh_token", row)
            self.assertIn("has_token", row)
            self.assertIn("token_tail", row)
            if row["has_token"]:
                self.assertTrue(row["token_tail"])
                self.assertLessEqual(len(row["token_tail"]), 6)

    def test_overview_stats_matches_legacy_counts(self) -> None:
        rows = account_store.list_accounts()
        expected_total = len(rows)
        expected_with_token = sum(1 for r in rows if r.get("has_token"))
        stats = account_store.overview_stats()
        self.assertEqual(stats["total"], expected_total)
        self.assertEqual(stats["with_token"], expected_with_token)
        self.assertIsInstance(stats["batches"], list)
        self.assertTrue(all({"label", "count"} <= set(b) for b in stats["batches"]))
        # 没批次的那条不该出现在批次分布里
        self.assertNotIn("", [b["label"] for b in stats["batches"]])


class AccountsApiTest(unittest.TestCase):
    """HTTP 层：分页响应形状 + 单号详情 + 轻量统计。"""

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
        for i in range(12):
            _seed_account(
                f"api{i:03d}@outlook.com",
                has_token=(i % 2 == 0),
                recovery=(i % 3 == 0),
                batch_no=1,
                batch_label="B1",
            )

    def test_default_limit_is_50_and_response_shape(self) -> None:
        r = self.client.get("/api/accounts")
        self.assertEqual(r.status_code, 200)
        body = r.json()
        self.assertEqual(body["limit"], 50)
        self.assertEqual(body["offset"], 0)
        self.assertEqual(body["total"], 12)
        self.assertEqual(len(body["accounts"]), 12)  # 12 < 50，一页装完
        self.assertIn("stats", body)
        self.assertEqual(body["stats"]["total"], 12)

    def test_limit_and_offset_paginate_correctly(self) -> None:
        r1 = self.client.get("/api/accounts", params={"limit": 5, "offset": 0})
        r2 = self.client.get("/api/accounts", params={"limit": 5, "offset": 5})
        r3 = self.client.get("/api/accounts", params={"limit": 5, "offset": 10})
        j1, j2, j3 = r1.json(), r2.json(), r3.json()
        self.assertEqual(len(j1["accounts"]), 5)
        self.assertEqual(len(j2["accounts"]), 5)
        self.assertEqual(len(j3["accounts"]), 2)
        emails = {a["email"] for a in j1["accounts"] + j2["accounts"] + j3["accounts"]}
        self.assertEqual(len(emails), 12)

    def test_list_rows_never_include_full_token_or_combo(self) -> None:
        r = self.client.get("/api/accounts")
        for row in r.json()["accounts"]:
            self.assertNotIn("refresh_token", row)
            self.assertNotIn("combo", row)
            self.assertNotIn("combo_dual", row)
            self.assertIn("has_token", row)
            self.assertIn("token_tail", row)

    def test_view_and_batch_and_q_query_params(self) -> None:
        r = self.client.get("/api/accounts", params={"view": "without"})
        self.assertTrue(all(not a["has_token"] for a in r.json()["accounts"]))

        r = self.client.get("/api/accounts", params={"batch": "B1"})
        self.assertEqual(r.json()["total"], 12)

        r = self.client.get("/api/accounts", params={"batch": "none"})
        self.assertEqual(r.json()["total"], 0)

        r = self.client.get("/api/accounts", params={"q": "api001"})
        self.assertEqual(r.json()["total"], 1)
        self.assertEqual(r.json()["accounts"][0]["email"], "api001@outlook.com")

    def test_account_detail_returns_full_secret_fields(self) -> None:
        r = self.client.get("/api/accounts/api000@outlook.com")
        self.assertEqual(r.status_code, 200)
        d = r.json()
        self.assertEqual(d["email"], "api000@outlook.com")
        self.assertTrue(d["refresh_token"].startswith("RT."))
        self.assertIn("----", d["combo"])
        self.assertTrue(d["has_token"])

    def test_account_detail_404_for_missing_email(self) -> None:
        r = self.client.get("/api/accounts/nobody@nowhere.com")
        self.assertEqual(r.status_code, 404)

    def test_accounts_stats_endpoint(self) -> None:
        r = self.client.get("/api/accounts/stats")
        self.assertEqual(r.status_code, 200)
        body = r.json()
        self.assertEqual(body["total"], 12)
        self.assertIn("recovery_pool_configured", body)
        self.assertIn("batches", body)

    def test_literal_account_routes_take_priority_over_email_route(self) -> None:
        """/api/accounts/export 等字面量路由要先于 /api/accounts/{email} 通配路由。"""
        r = self.client.get("/api/accounts/export")
        self.assertEqual(r.status_code, 200)
        self.assertEqual(r.headers["content-type"].split(";")[0], "text/plain")


class AccountsPayloadSizeRegressionTest(unittest.TestCase):
    """体量回归：模拟"近万条大 token 账号"场景的缩小版，锁住分页返回体积。"""

    client: TestClient
    N = 800

    @classmethod
    def setUpClass(cls) -> None:
        from webapp import server

        cls._ctx = TestClient(server.app)
        cls.client = cls._ctx.__enter__()
        _clear_accounts()
        for i in range(cls.N):
            _seed_account(
                f"bulk{i:05d}@outlook.com",
                has_token=True,
                token_len=1400,  # 真实 Graph refresh_token 常见长度量级
                recovery=True,
                batch_no=(i % 30) + 1,
                batch_label=f"B{(i % 30) + 1}",
            )

    @classmethod
    def tearDownClass(cls) -> None:
        cls._ctx.__exit__(None, None, None)
        _clear_accounts()

    def test_full_unpaginated_payload_would_be_multi_megabyte(self) -> None:
        """先证明「不分页」确实是 MB 级——这正是修复前的 bug 现场。"""
        rows = account_store.list_accounts()
        size = len(json.dumps(rows, ensure_ascii=False).encode("utf-8"))
        self.assertEqual(len(rows), self.N)
        self.assertGreater(size, 1_000_000, "全量账号 JSON 应在 MB 级，不然这个回归测试没意义")

    def test_default_page_response_is_small_and_fast_enough(self) -> None:
        r = self.client.get("/api/accounts")  # 默认 limit=50
        self.assertEqual(r.status_code, 200)
        body = r.json()
        self.assertEqual(body["total"], self.N)
        self.assertEqual(len(body["accounts"]), 50)
        # 验收标准：limit=50 响应 < 500KB（这里留足余量判定为 200KB，防止字段膨胀）
        self.assertLess(len(r.content), 200 * 1024)

    def test_stats_endpoint_stays_small_regardless_of_pool_size(self) -> None:
        r = self.client.get("/api/accounts/stats")
        self.assertEqual(r.status_code, 200)
        self.assertEqual(r.json()["total"], self.N)
        self.assertLess(len(r.content), 20 * 1024)


if __name__ == "__main__":
    unittest.main()
