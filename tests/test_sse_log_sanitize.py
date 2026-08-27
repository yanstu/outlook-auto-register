"""SSE 注册日志净化：丢弃内部调试行，替换开发腔。"""
from __future__ import annotations

import logging
from unittest import mock

from webapp.server import _JobLogHandler, _sanitize_sse_log


def test_discard_internal_debug_lines():
    samples = [
        "已 dump proofs_add_resp.html (1234 bytes)",
        "credentialaction 构造 POST proofs/Add（对齐 Outlook抓包.har idx 100）",
        "proofs 页未解析到 frmAddProof 表单",
        "VerifyProof 提交 iOttText=123456 proof=x canary=abc…",
        "$Config.urlPost → https://login.live.com/xx",
        "✅ 产出自检：令牌可读信 resource=graph /me=200 name=Ada",
        "会话就绪 uaid=abc123 opid=xyz",
        "纯协议 verify #2 attempt=1 state=continue",
        "token 交换响应非 JSON status=400 body={\"error\":1}",
        "比特收割成功 px3=aaaa... pressed=True",
        "任务超时 task_id=abc",
        "slt 登录完成 status=200 url=https://x proofs=False",
        "Passkey 按 HAR 取消报名 enroll error_code=NotAllowedError",
        "Passkey fido/create 表单无法取消",
    ]
    for msg in samples:
        assert _sanitize_sse_log(msg) is None, msg


def test_replace_dev_jargon():
    cases = [
        ("提交 OAuth 同意", "OAuth", "授权"),
        ("已获取 refresh_token", "refresh_token", "读信令牌"),
        ("账号已写入 SQLite", "SQLite", "账号库"),
        ("proofs 收码池：选用恢复老号 a***", "proofs", "安全验证"),
        ("AddProof 成功", "AddProof", "绑定恢复邮箱"),
        ("VerifyProof 成功", "VerifyProof", "验证恢复邮箱"),
        ("risk/initialize 完成", "risk/initialize", "初始化"),
        ("risk/verify 进行中", "risk/verify", "人机验证"),
        ("恢复老号令牌失效(invalid_grant)", "invalid_grant", "令牌失效"),
        ("已从恢复老号读到 OTT", "OTT", "验证码"),
        ("captcha.run 已建 task", "captcha.run", "打码服务"),
        ("coolhs-mail 分配地址失败", "coolhs-mail", "恢复邮箱"),
        ("防封·一号一 IP：已分配", "防封·一号一 IP", "独立出口"),
        ("防封·启动错峰：间隔 3 秒", "防封·启动错峰", "启动间隔"),
        ("mail OAuth 授权…", "mail OAuth", "读信授权"),
        ("mail OAuth 授权…", "OAuth", "读信授权"),
        ("token 交换成功", "token 交换", "换取令牌"),
        ("阶段耗时(s): {'a': 1}", "阶段耗时(s):", "各阶段耗时："),
    ]
    for src, old, new in cases:
        out = _sanitize_sse_log(src)
        assert out is not None, src
        assert old not in out, (src, out, old)
        assert new in out, (src, out, new)


def test_chain_and_keep_business_lines():
    assert _sanitize_sse_log("选用邮箱: ada@outlook.com") == "选用邮箱: ada@outlook.com"
    assert _sanitize_sse_log("账号已保存（孵化中）: ada@outlook.com") == "账号已保存（孵化中）: ada@outlook.com"
    assert _sanitize_sse_log("人机验证通过") == "人机验证通过"
    assert _sanitize_sse_log("换取令牌成功，已获取读信令牌") == "换取令牌成功，已获取读信令牌"
    assert _sanitize_sse_log("") is None
    assert _sanitize_sse_log("   ") is None


def test_long_http_url_truncated():
    long_url = "https://login.live.com/oauth20_authorize.srf?client_id=abc&scope=mail"
    assert len(long_url) > 40
    out = _sanitize_sse_log(f"跳过账号安全信息插页 → {long_url}")
    assert out is not None
    assert "http" not in out
    assert "…" in out
    short = "http://a.co"
    assert _sanitize_sse_log(f"见 {short}") == f"见 {short}"


def test_engine_hot_path_samples_have_no_jargon():
    """源头改写 + sanitizer 之后，用户不该再看到这些词。"""
    leftover = [
        "获取授权登录页…",
        "会话就绪",
        "预加载安全组件…",
        "各阶段耗时：{'pick_email': 0.2}",
        "✅ 产出自检：令牌可读信 name=Ada",
        "独立出口：已为 3 个账号各分配独立出口。",
        "启动间隔：相邻账号注册随机间隔 3.0–8.0 秒启动。",
        "人机验证通过",
        "初始化完成",
        "提交授权同意",
        "读信授权（consumer）…",
        "换取令牌成功，已获取读信令牌",
    ]
    banned = ("OAuth", "HAR", "SQLite", "refresh_token", "uaid=", "opid=", "/me=")
    for msg in leftover:
        clean = _sanitize_sse_log(msg)
        assert clean is not None, msg
        for word in banned:
            assert word not in clean, (msg, clean, word)


def test_job_handler_drops_and_rewrites():
    job = mock.Mock()
    handler = _JobLogHandler()
    rec_drop = logging.LogRecord(
        "outlook_api_reg.post_register", logging.INFO, __file__, 1,
        "已 dump proofs_add_resp.html (12 bytes)", (), None,
    )
    rec_keep = logging.LogRecord(
        "outlook_api_reg.account_store", logging.INFO, __file__, 1,
        "账号已写入 SQLite: ada@outlook.com", (), None,
    )
    with mock.patch("webapp.server._jobs", {"job-1": job}), \
            mock.patch("webapp.server._active_batch_job", "job-1"):
        handler.emit(rec_drop)
        handler.emit(rec_keep)
    calls = [c.args for c in job.push_log.call_args_list]
    assert len(calls) == 1
    assert calls[0][0] == "INFO"
    assert "SQLite" not in calls[0][1]
    assert "账号库" in calls[0][1]
    assert "ada@outlook.com" in calls[0][1]
