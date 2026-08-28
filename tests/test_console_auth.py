"""运维台登录门：cookie 会话 / 口令头 / 公网例外，以及 Mailbox API 的独立通道。"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from fastapi.testclient import TestClient

ROOT = Path(__file__).resolve().parents[1]
UNIT_TEMPLATE = ROOT / "deploy" / "outlook-console.user.service"
DEPLOY_SCRIPT = ROOT / "scripts" / "deploy_to_38.sh"
PASSWORD_PLACEHOLDER = "__OUTLOOK_CONSOLE_PASSWORD__"

PASSWORD = "s3cr3t-console"
ADMIN_KEY = "console-auth-admin-key"
ADMIN_HEADERS = {"Authorization": f"Bearer {ADMIN_KEY}"}

# 同一次 pytest 里还有别的用例在用它们自己的库和 admin key。环境变量都是进程级的，
# 所以这里只在用例存活期间打补丁，退出时原样还回去——否则会把兄弟用例的库指到这儿，
# 或者把开发机上真正的 accounts/outlook.db 当成测试库写进去。
_TMP_DIR = tempfile.mkdtemp(prefix="console-auth-test-")
_ENV_PATCH = {
    "OUTLOOK_DB_PATH": os.path.join(_TMP_DIR, "outlook.db"),
    "OUTLOOK_MAILBOX_API_ENABLED": "1",
    "OUTLOOK_MAILBOX_API_ADMIN_KEY": ADMIN_KEY,
}


def tearDownModule() -> None:
    shutil.rmtree(_TMP_DIR, ignore_errors=True)


class ConsoleAuthTest(unittest.TestCase):
    """口令非空：除探活 / 登录页 / Mailbox API 外，一律要会话。"""

    client: TestClient

    @classmethod
    def setUpClass(cls) -> None:
        cls._env = mock.patch.dict(os.environ, {**_ENV_PATCH, "OUTLOOK_CONSOLE_PASSWORD": PASSWORD})
        cls._env.start()
        from webapp import server

        cls.server = server
        cls._ctx = TestClient(server.app)
        cls.client = cls._ctx.__enter__()

    @classmethod
    def tearDownClass(cls) -> None:
        cls._ctx.__exit__(None, None, None)
        cls._env.stop()

    def setUp(self) -> None:
        # 每个用例从「未登录」开始：上一个用例种下的 cookie 不能漏过来
        self.client.cookies.clear()
        os.environ["OUTLOOK_CONSOLE_PASSWORD"] = PASSWORD

    # ---- 公开面 ----------------------------------------------------------
    def test_health_is_public_and_advertises_the_gate(self):
        r = self.client.get("/api/health")
        self.assertEqual(r.status_code, 200)
        self.assertTrue(r.json()["auth_required"])

    def test_login_page_is_public(self):
        r = self.client.get("/login.html")
        self.assertEqual(r.status_code, 200)
        self.assertIn("Outlook 注册控制台", r.text)

    # ---- 未登录 ----------------------------------------------------------
    def test_console_page_redirects_to_login(self):
        r = self.client.get("/", follow_redirects=False)
        self.assertEqual(r.status_code, 302)
        self.assertEqual(r.headers["location"], "/login.html?next=%2F")

    def test_api_without_session_is_401_json(self):
        r = self.client.get("/api/config")
        self.assertEqual(r.status_code, 401)
        self.assertTrue(r.json()["auth"])

    def test_non_api_non_page_path_is_401_text(self):
        r = self.client.get("/openapi.json")
        self.assertEqual(r.status_code, 401)

    def test_wrong_password_header_rejected(self):
        r = self.client.get("/api/config", headers={"X-Console-Password": "nope"})
        self.assertEqual(r.status_code, 401)

    def test_forged_session_cookie_rejected(self):
        self.client.cookies.set(self.server.CONSOLE_SESSION_COOKIE, "deadbeef")
        r = self.client.get("/api/config")
        self.assertEqual(r.status_code, 401)

    def test_cookie_minted_for_another_password_rejected(self):
        stale = self.server._console_session_token("old-password")
        self.client.cookies.set(self.server.CONSOLE_SESSION_COOKIE, stale)
        r = self.client.get("/api/config")
        self.assertEqual(r.status_code, 401)

    # ---- 登录 ------------------------------------------------------------
    def test_login_wrong_password(self):
        r = self.client.post("/api/auth/login", json={"password": "wrong"})
        self.assertEqual(r.status_code, 401)
        self.assertNotIn(self.server.CONSOLE_SESSION_COOKIE, self.client.cookies)

    def test_login_then_browse(self):
        r = self.client.post("/api/auth/login", json={"password": PASSWORD})
        self.assertEqual(r.status_code, 200)
        self.assertIn(self.server.CONSOLE_SESSION_COOKIE, self.client.cookies)

        page = self.client.get("/", follow_redirects=False)
        self.assertEqual(page.status_code, 200)
        self.assertEqual(self.client.get("/api/config").status_code, 200)
        self.assertTrue(self.client.get("/api/auth/check").json()["auth_required"])

    def test_logout_drops_the_session(self):
        self.client.post("/api/auth/login", json={"password": PASSWORD})
        self.assertEqual(self.client.get("/api/config").status_code, 200)
        self.client.post("/api/auth/logout")
        self.assertEqual(self.client.get("/api/config").status_code, 401)

    def test_header_password_works_and_mints_a_cookie(self):
        r = self.client.get("/api/config", headers={"X-Console-Password": PASSWORD})
        self.assertEqual(r.status_code, 200)
        self.assertIn(self.server.CONSOLE_SESSION_COOKIE, self.client.cookies)
        # 拿到 cookie 后不带头也能继续
        self.assertEqual(self.client.get("/api/config").status_code, 200)

    def test_session_dies_when_password_changes(self):
        self.client.post("/api/auth/login", json={"password": PASSWORD})
        self.assertEqual(self.client.get("/api/config").status_code, 200)
        os.environ["OUTLOOK_CONSOLE_PASSWORD"] = "rotated-password"
        self.assertEqual(self.client.get("/api/config").status_code, 401)

    # ---- Mailbox API 不受运维口令影响 --------------------------------------
    def test_mailbox_api_ignores_console_password(self):
        r = self.client.get("/api/v1/health", headers=ADMIN_HEADERS)
        self.assertEqual(r.status_code, 200)
        self.assertTrue(r.json()["ok"])

    def test_mailbox_api_still_needs_its_own_bearer(self):
        r = self.client.get("/api/v1/health")
        self.assertEqual(r.status_code, 401)
        # 运维台的 401 长 {"ok":false,"error":...}，Mailbox API 的长 {"code":...}
        self.assertEqual(r.json()["code"], "unauthorized")

    def test_console_session_does_not_unlock_mailbox_api(self):
        self.client.post("/api/auth/login", json={"password": PASSWORD})
        r = self.client.get("/api/v1/health")
        self.assertEqual(r.status_code, 401)
        self.assertEqual(r.json()["code"], "unauthorized")


class ConsoleAuthDisabledTest(unittest.TestCase):
    """口令留空（本地开发）：不拦任何请求，也不假装有登录态。"""

    client: TestClient

    @classmethod
    def setUpClass(cls) -> None:
        cls._env = mock.patch.dict(os.environ, _ENV_PATCH)
        cls._env.start()
        os.environ.pop("OUTLOOK_CONSOLE_PASSWORD", None)
        from webapp import server

        cls.server = server
        cls._ctx = TestClient(server.app)
        cls.client = cls._ctx.__enter__()

    @classmethod
    def tearDownClass(cls) -> None:
        cls._ctx.__exit__(None, None, None)
        cls._env.stop()

    def test_everything_open(self):
        self.assertEqual(self.client.get("/", follow_redirects=False).status_code, 200)
        self.assertEqual(self.client.get("/api/config").status_code, 200)
        self.assertFalse(self.client.get("/api/health").json()["auth_required"])

    def test_login_is_a_noop(self):
        r = self.client.post("/api/auth/login", json={"password": ""})
        self.assertEqual(r.status_code, 200)
        self.assertFalse(r.json()["auth_required"])

    def test_mailbox_api_still_guarded(self):
        self.assertEqual(self.client.get("/api/v1/health").status_code, 401)


class WhitespacePasswordTest(unittest.TestCase):
    """只有空白字符的口令等于没设，不能让它伪装成一道门。"""

    def test_blank_password_is_no_password(self):
        from webapp import server

        prev = os.environ.get("OUTLOOK_CONSOLE_PASSWORD")
        os.environ["OUTLOOK_CONSOLE_PASSWORD"] = "   "
        try:
            self.assertEqual(server._console_password(), "")
        finally:
            if prev is None:
                os.environ.pop("OUTLOOK_CONSOLE_PASSWORD", None)
            else:
                os.environ["OUTLOOK_CONSOLE_PASSWORD"] = prev


class DeployedPasswordIsLiteralTest(unittest.TestCase):
    """部署链路必须把口令原样搬到 systemd 单元里，一个字符都不许变形。

    线上那枚口令以标点结尾（真值只在 ``.deploy/local.env``），而它要穿过两层：
    deploy_to_38.sh 的占位符替换，和 systemd 的 ``Environment=`` 解析。任何一层把
    尾部的点号、引号或空白吃掉或加上，运维台就会变成一个谁都登不进去的空壳。
    """

    # 尾随点号、引号、空白、`$` —— 把 systemd 与 shell 各自的敏感字符凑齐
    SAMPLES = ("sample1...", "a.b.c", "p@ss w0rd", "with'quote\"s", "tail$dollar")

    def _render(self, password: str) -> str:
        return UNIT_TEMPLATE.read_text(encoding="utf-8").replace(PASSWORD_PLACEHOLDER, password)

    def test_template_and_script_agree_on_the_placeholder(self):
        # 占位符在两边改名不同步的话，部署会把 "__OUTLOOK_CONSOLE_PASSWORD__"
        # 这串字面量当成口令发上线，而且探活全绿、毫无察觉。
        self.assertIn(PASSWORD_PLACEHOLDER, UNIT_TEMPLATE.read_text(encoding="utf-8"))
        self.assertIn(PASSWORD_PLACEHOLDER, DEPLOY_SCRIPT.read_text(encoding="utf-8"))

    def test_rendered_unit_carries_the_password_verbatim(self):
        for password in self.SAMPLES:
            with self.subTest(password=password):
                rendered = self._render(password)
                self.assertNotIn(PASSWORD_PLACEHOLDER, rendered)
                line = re.search(
                    r"^Environment=OUTLOOK_CONSOLE_PASSWORD=(.*)$", rendered, re.M
                )
                self.assertIsNotNone(line)
                self.assertEqual(line.group(1), password)

    def test_server_reads_back_a_punctuated_password(self):
        from webapp import server

        for password in self.SAMPLES:
            with self.subTest(password=password):
                with mock.patch.dict(os.environ, {"OUTLOOK_CONSOLE_PASSWORD": password}):
                    self.assertEqual(server._console_password(), password)


if __name__ == "__main__":
    unittest.main()
