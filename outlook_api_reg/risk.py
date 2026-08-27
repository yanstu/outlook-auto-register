from __future__ import annotations

import logging
import os
import sys
import time
from typing import Any, Optional

import requests

from .api import build_msa_risk_verify_signature, risk_initialize, risk_verify
from .bootstrap import preload_px_challenge_assets
from .captcha import CaptchaRunTask, create_captcha_run_task, poll_captcha_run_token, solve_perimeterx
from .http_session import OutlookHttpSession
from .models import AccountInfo, SignupSession
from .px_collector import load_challenge_iframe, post_px_beacon, post_px_bundle, warmup_px_session
from .px_cookies import (
    bind_press_solution,
    build_challenge_solution,
    build_px_metadata,
    solver_context,
)

logger = logging.getLogger(__name__)


class RegisterRetryable(RuntimeError):
    """可通过换新住宅 IP 重试整个注册的错误基类。"""


class Verify2Failed(RegisterRetryable):
    """verify #2（PX 按住）未通过。"""


class RiskBlocked(RegisterRetryable):
    """verify #1 riskBlock（AADSTS7005106，该 IP 被风控拦截）。"""


def load_human_sensor(http: OutlookHttpSession, ctx: SignupSession) -> None:
    url = ctx.human_sensor_url
    if not url:
        return
    logger.info("加载安全检测…")
    http.get(url, headers={"Referer": ctx.signup_page_url})


def _solver_ctx(
    http: OutlookHttpSession,
    ctx: SignupSession,
    *,
    proxy: Optional[str],
    challenge_meta: Optional[dict[str, Any]] = None,
    country: str = "US",
) -> dict[str, Any]:
    return solver_context(
        http.session,
        page_url=ctx.signup_page_url,
        uaid=ctx.uaid,
        challenge_meta=challenge_meta or ctx.px_challenge_meta,
        proxy=proxy or http.proxy,
        country=country,
    )


def _ensure_captcha_run_silent_task(
    http: OutlookHttpSession,
    ctx: SignupSession,
    *,
    proxy: Optional[str],
    country: str = "US",
) -> Optional[CaptchaRunTask]:
    existing = ctx.captcha_run_task
    if isinstance(existing, CaptchaRunTask):
        return existing
    sctx = _solver_ctx(http, ctx, proxy=proxy, challenge_meta=None, country=country)
    task = create_captcha_run_task(sctx)
    if task:
        ctx.captcha_run_task = task
    return task


def _poll_captcha_run_press(
    http: OutlookHttpSession,
    ctx: SignupSession,
    *,
    challenge_meta: dict[str, Any],
) -> dict[str, str]:
    """官方文档：verify#1 同一 taskId 上 GET press（须已 GET silent）。"""
    task = ctx.captcha_run_task
    if not isinstance(task, CaptchaRunTask):
        raise RuntimeError("captcha.run 须先在 verify#1 前 POST 建 task 并 GET silent")

    stable_vid = str(challenge_meta.get("vid", ""))
    if not task.silent_fetched:
        logger.debug("captcha.run press 前补拉 silent（同 task=%s）", task.task_id)
        poll_captcha_run_token(task, "silent")

    logger.debug(
        "captcha.run GET press task=%s challenge_uuid=%s vid=%s",
        task.task_id,
        challenge_meta.get("uuid", ""),
        challenge_meta.get("vid", ""),
    )
    solved = poll_captcha_run_token(task, "press", warmup_silent_before_press=False)
    if not solved or not solved.get("px3"):
        raise RuntimeError(
            f"captcha.run press 未返回 pressToken（task={task.task_id}），"
            "请核对代理与官方文档：同 task 先 silent 后 press"
        )
    px3 = solved.get("px3", "")
    if ":1000:" not in px3:
        logger.debug("pressToken px3 无 :1000: 段（HAR verify#2 成功样本均含 :1000:）")
    return http.apply_px_tokens(solved, preserve_vid=stable_vid)


def _solve_via_bitbrowser(
    http: OutlookHttpSession,
    *,
    phase: str,
    proxy: Optional[str],
    stable_vid: str = "",
) -> dict[str, str]:
    """用比特指纹浏览器收割 _px3（与注册同 IP），替代 captcha.run。"""
    _px_dir = os.path.join(os.path.dirname(__file__), "..", "px_solver")
    if _px_dir not in sys.path:
        sys.path.insert(0, _px_dir)
    from bit_px_solver import harvest  # noqa: E402

    p = proxy or http.proxy
    logger.debug("PX 走比特浏览器收割 phase=%s（同 IP=%s）", phase, str(p)[:40])
    sol = harvest(p, want_press=(phase == "press"))
    if not sol.get("px3"):
        raise RuntimeError(f"比特浏览器未收割到 _px3 phase={phase}")
    logger.debug("比特收割成功 px3=%s... pressed=%s", sol["px3"][:30], sol.get("pressed"))
    return http.apply_px_tokens(
        {"px3": sol["px3"], "pxde": sol.get("pxde", ""), "pxvid": sol.get("pxvid", "")},
        preserve_vid=stable_vid or sol.get("pxvid", ""),
    )


def _solve_px_protocol(
    http: OutlookHttpSession,
    ctx: SignupSession,
    *,
    phase: str,
    proxy: Optional[str],
    challenge_meta: Optional[dict[str, Any]] = None,
    country: str = "US",
) -> dict[str, str]:
    prefer = "silent" if phase == "silent" else "press"
    logger.info("人机验证进行中")
    logger.debug("纯协议打码 phase=%s prefer=%s", phase, prefer)

    if os.environ.get("PX_SOLVER", "").strip().lower() == "bitbrowser":
        _meta = challenge_meta or ctx.px_challenge_meta
        _vid = str(_meta.get("vid", "")) if phase == "press" else ""
        return _solve_via_bitbrowser(http, phase=phase, proxy=proxy, stable_vid=_vid)

    meta = challenge_meta or ctx.px_challenge_meta
    stable_vid = str(meta.get("vid", "")) if phase == "press" else ""
    sctx = _solver_ctx(http, ctx, proxy=proxy, challenge_meta=meta, country=country)

    if phase == "silent":
        task = _ensure_captcha_run_silent_task(
            http, ctx, proxy=proxy, country=country,
        )
        if task:
            solved = poll_captcha_run_token(task, "silent")
            if solved and solved.get("px3"):
                logger.info("人机验证通过")
                logger.debug("captcha.run silent 成功（task=%s）", task.task_id)
                return http.apply_px_tokens(solved, preserve_vid=stable_vid)
    elif phase == "press":
        try:
            return _poll_captcha_run_press(
                http, ctx, challenge_meta=meta,
            )
        except RuntimeError as exc:
            fallback = os.environ.get("PX_PRESS_FALLBACK", "").strip().lower()
            if fallback not in {"1", "true", "ez", "ezcaptcha", "capsolver", "auto"}:
                raise
            logger.warning("打码服务失败，改用备用方案")
            logger.debug("captcha.run press 失败，走 PX_PRESS_FALLBACK=%s: %s", fallback, exc)

    if phase == "silent":
        solved = solve_perimeterx(sctx, prefer_mode="silent")
    else:
        solved = solve_perimeterx(sctx, prefer_mode="press")
        if not solved or not solved.get("px3"):
            raise RuntimeError(
                "press 打码失败（captcha.run + fallback）。"
                "设 PX_PRESS_FALLBACK=ezcaptcha 可启用 EzCaptcha/CapSolver"
            )
    if not solved or not solved.get("px3"):
        raise RuntimeError(f"纯协议打码失败 phase={phase}，请检查 CAPTCHA_RUN_API_KEY / 代理 / 余额")
    return http.apply_px_tokens(solved, preserve_vid=stable_vid)


def _acquire_silent_px(
    http: OutlookHttpSession,
    ctx: SignupSession,
    *,
    mode: str,
    proxy: Optional[str],
    country: str,
    force_fresh: bool = False,
) -> dict[str, str]:
    """risk/verify #1 的 silent px：纯协议 captcha.run silent。"""
    px = http.px_cookies()
    if px.get("px3") and not force_fresh:
        logger.debug("silent px 复用已有 cookie")
        return px

    return _solve_px_protocol(http, ctx, phase="silent", proxy=proxy, country=country)


def _verify2(
    http: OutlookHttpSession,
    ctx: SignupSession,
    *,
    challenge_meta: dict[str, Any],
    px: dict[str, str],
    challenge_type: str,
) -> dict[str, Any]:
    bound = bind_press_solution(px, challenge_meta)
    return risk_verify(
        http, ctx,
        continuation_token=ctx.continuation_token,
        risk_provider_metadata=build_px_metadata(bound),
        challenge_solution=build_challenge_solution(
            px, challenge_meta, challenge_type=challenge_type,
        ),
    )


def _protocol_verify2(
    http: OutlookHttpSession,
    ctx: SignupSession,
    account: AccountInfo,
    *,
    proxy: Optional[str],
    challenge_meta: dict[str, Any],
    challenge_type: str,
) -> bool:
    """纯协议 verify #2：复用 verify#1 同一 captcha.run task 拉 press。

    captcha.run 单 task = 单 PX 会话；同一挑战重复取 press 会拿到缓存的同一 token，
    被微软拒过就会再被拒，多轮无意义。故只解 1 次，失败交由上层换新住宅 IP 重试。
    """
    max_attempts = 1
    meta = challenge_meta
    for attempt in range(1, max_attempts + 1):
        try:
            challenge_px = _solve_px_protocol(
                http, ctx, phase="press", proxy=proxy, challenge_meta=meta,
                country=account.country,
            )
            resp2 = _verify2_with_retry(
                http, ctx,
                challenge_meta=meta,
                px=challenge_px,
                challenge_type=challenge_type,
            )
            state = resp2.get("state", "")
            logger.debug("纯协议 verify #2 attempt=%s state=%s", attempt, state)
            if state == "continue":
                logger.info("人机验证通过")
                return True
            logger.debug("verify #2 body keys=%s", list(resp2.keys()))
            nxt = resp2.get("challengeDetails", {}).get("challengeMetadata", {})
            if nxt:
                meta = nxt
                ctx.px_challenge_meta = meta
        except RuntimeError as exc:
            logger.warning("人机验证打码失败: %s", exc)
        except requests.HTTPError as exc:
            if exc.response is not None and exc.response.status_code == 403:
                logger.error("人机验证被拦截，将换出口重试")
                return False
            raise
        except (requests.Timeout, requests.ConnectionError) as exc:
            logger.warning("人机验证网络异常: %s", exc)
        time.sleep(2.0)
    return False


def _verify2_with_retry(
    http: OutlookHttpSession,
    ctx: SignupSession,
    *,
    challenge_meta: dict[str, Any],
    px: dict[str, str],
    challenge_type: str,
    retries: int = 3,
) -> dict[str, Any]:
    last_exc: Optional[Exception] = None
    for i in range(1, retries + 1):
        try:
            return _verify2(
                http, ctx,
                challenge_meta=challenge_meta,
                px=px,
                challenge_type=challenge_type,
            )
        except (requests.Timeout, requests.ConnectionError) as exc:
            last_exc = exc
            logger.warning("人机验证请求超时，重试 %s/%s: %s", i, retries, exc)
            time.sleep(3.0)
    if last_exc:
        raise last_exc
    raise RuntimeError("verify #2 重试耗尽")


def solve_risk_challenge(
    http: OutlookHttpSession,
    ctx: SignupSession,
    account: AccountInfo,
    *,
    mode: str = "solver",
    proxy: Optional[str] = None,
) -> None:
    """风控链（纯协议 only）：silent 打码 → verify #1 → press 打码 → verify #2。"""
    init = risk_initialize(http, ctx, "")
    logger.info("初始化完成")
    logger.debug("risk/initialize state=%s", init.get("state"))

    if ctx.human_sensor_url:
        load_human_sensor(http, ctx)

    if not ctx.continuation_token:
        raise RuntimeError("risk/initialize 未返回 continuationToken")

    px_meta = _acquire_silent_px(http, ctx, mode=mode, proxy=proxy, country=account.country)

    signature = build_msa_risk_verify_signature(account, ctx)
    try:
        resp1 = risk_verify(
            http, ctx,
            continuation_token=ctx.continuation_token,
            risk_provider_metadata=build_px_metadata(px_meta),
            msa_risk_verify_signature=signature,
        )
    except requests.HTTPError as exc:
        if exc.response is not None and exc.response.status_code == 403:
            body = ""
            try:
                body = exc.response.text[:200]
            except Exception:  # noqa: BLE001
                pass
            # riskBlock：该住宅 IP 被微软拦截，同 IP 重试无意义，换新 IP 重试整个注册
            raise RiskBlocked(f"verify #1 riskBlock（该 IP 被拦，换新 IP 重试）: {body}") from exc
        raise

    state = resp1.get("state", "")
    logger.debug("risk/verify #1 state=%s", state)
    if state == "continue":
        logger.info("人机验证通过")
        return
    if state != "riskChallengeRequired":
        raise RuntimeError(f"risk/verify #1 未预期状态: {state}")

    challenge = resp1.get("challengeDetails", {})
    challenge_meta = challenge.get("challengeMetadata", {}) or {}
    challenge_type = challenge.get("challengeType", "HumanCaptcha")
    ctx.px_challenge_meta = challenge_meta

    load_challenge_iframe(http, ctx, challenge_meta)
    preload_px_challenge_assets(http, ctx, challenge_meta)
    warmup_px_session(http, ctx)
    time.sleep(1.0)

    post_px_beacon(http, ctx, tag="pre-press")
    post_px_bundle(http, ctx, tag="pre-press")

    if _protocol_verify2(
        http, ctx, account,
        proxy=proxy,
        challenge_meta=challenge_meta,
        challenge_type=challenge_type,
    ):
        return
    raise Verify2Failed("risk/verify #2 未通过（纯协议 captcha.run press）")
