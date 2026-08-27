from __future__ import annotations

import hashlib
import json
import logging
import re
import secrets
import urllib.parse
from base64 import urlsafe_b64encode
from typing import Any, Optional

from .constants import (
    COBRAND_ID,
    DEFAULT_LC,
    DEFAULT_MKT,
    OUTLOOK_CLIENT_ID,
    OUTLOOK_REDIRECT_URI,
    OUTLOOK_SCOPE,
    PX_APP_ID,
    PX_COLLECTOR_BASE,
)
from .http_session import OutlookHttpSession
from .models import SignupSession

logger = logging.getLogger(__name__)


def _b64url(data: bytes) -> str:
    return urlsafe_b64encode(data).rstrip(b"=").decode()


def _parse_server_data(html: str) -> dict:
    marker = html.find("var ServerData=")
    if marker < 0:
        marker = html.find("var ServerData =")
    if marker < 0:
        raise ValueError("signup 页面未找到 ServerData")
    json_start = html.find("{", marker)
    obj, _ = json.JSONDecoder().raw_decode(html, json_start)
    return obj


def _pkce_pair() -> tuple[str, str]:
    verifier = secrets.token_urlsafe(64)[:128]
    challenge = _b64url(hashlib.sha256(verifier.encode()).digest())
    return verifier, challenge


def bootstrap_session(http: OutlookHttpSession, *, mkt: str = DEFAULT_MKT, lc: str = DEFAULT_LC) -> SignupSession:
    code_verifier, code_challenge = _pkce_pair()

    auth_qs = urllib.parse.urlencode({
        "client_id": OUTLOOK_CLIENT_ID,
        "cobrandid": COBRAND_ID,
        "response_type": "code",
        "redirect_uri": OUTLOOK_REDIRECT_URI,
        "scope": OUTLOOK_SCOPE,
        "code_challenge": code_challenge,
        "code_challenge_method": "S256",
        "mkt": mkt,
        "lc": lc,
    })
    auth_url = f"https://login.live.com/oauth20_authorize.srf?{auth_qs}"
    logger.info("获取授权登录页…")
    login_resp = http.get(auth_url, allow_redirects=True)
    login_resp.raise_for_status()

    m = re.search(r'"(https://signup\.live\.com/signup[^"]+)"', login_resp.text)
    if not m:
        raise RuntimeError("登录页未找到注册入口，授权参数可能有误")

    signup_link = m.group(1).replace("\\u0026", "&").replace("\\/", "/")
    logger.info("加载注册页…")
    signup_resp = http.get(signup_link, allow_redirects=True)
    signup_resp.raise_for_status()

    sd = _parse_server_data(signup_resp.text)
    parsed = urllib.parse.urlparse(signup_link)
    qs = urllib.parse.parse_qs(parsed.query)

    uaid_m = re.search(r"uaid=([a-f0-9]{32})", signup_link)
    uaid = uaid_m.group(1) if uaid_m else qs.get("uaid", [""])[0]
    if not uaid:
        raise RuntimeError("未解析到 uaid")

    ctx = SignupSession(
        uaid=uaid,
        signup_url=signup_link,
        signup_page_url=signup_resp.url,
        cobrandid=qs.get("cobrandid", [COBRAND_ID])[0],
        contextid=qs.get("contextid", [""])[0],
        opid=qs.get("opid", [""])[0],
        bk=qs.get("bk", [""])[0],
        sru=qs.get("sru", [""])[0],
        canary=sd.get("apiCanary", ""),
        hpgid=int(sd.get("hpgid", 200225)),
        scid=int(sd.get("iScenarioId", 100118)),
        server_data=sd,
        code_verifier=code_verifier,
        code_challenge=code_challenge,
        mkt=mkt,
        lc=lc,
    )
    if not ctx.canary:
        raise RuntimeError("ServerData 中无 apiCanary")

    http.signup_ctx = ctx
    logger.info("会话就绪")
    return ctx


def preload_perimeterx(http: OutlookHttpSession, ctx: SignupSession) -> None:
    """预加载 PerimeterX DFP + iframe + collector（对齐 HAR）。"""
    captcha_info = ctx.server_data.get("oCaptchaInfo", {})
    dfp_url = captcha_info.get("urlDfp")
    human_url = captcha_info.get("urlHumanIframe")

    urls = [
        dfp_url,
        human_url,
        f"{PX_COLLECTOR_BASE}/api/v2/msft/beacon",
        f"{PX_COLLECTOR_BASE}/assets/js/bundle",
    ]
    logger.info("预加载安全组件…")
    for url in urls:
        if not url:
            continue
        logger.debug("PX 预加载: %s", url[:100])
        try:
            http.get(url, headers={"Referer": ctx.signup_page_url})
        except Exception as exc:
            logger.debug("PX 预加载跳过 %s: %s", url[:60], exc)

    logger.debug("PX 预加载后 cookies: %s", http.cookie_names())


def preload_px_challenge_assets(
    http: OutlookHttpSession,
    ctx: SignupSession,
    challenge_meta: dict[str, Any],
) -> None:
    """riskChallengeRequired 后加载 challenge 相关资源。"""
    uuid = challenge_meta.get("uuid", "")
    vid = challenge_meta.get("vid", "")
    challenge_url = challenge_meta.get("challengeUrl", "")

    urls = [
        challenge_url,
        f"https://stk.hsprotect.net/ns?c={uuid}" if uuid else "",
        (
            f"https://captcha.hsprotect.net/{PX_APP_ID}/captcha.js"
            f"?a=c&m=0&u={uuid}&v={vid}"
            if uuid and vid else ""
        ),
        f"https://df.cfp.microsoft.com/Clear.HTML?ctx=Ls1.0&wl=False&session_id={uuid}" if uuid else "",
        f"https://fpt.live.com/Images/Clear.PNG?ctx=jscb1.0&session_id={uuid}" if uuid else "",
    ]
    for url in urls:
        if not url:
            continue
        logger.debug("挑战资源: %s", url[:100])
        try:
            http.get(url, headers={"Referer": ctx.signup_page_url})
        except Exception as exc:
            logger.debug("挑战资源跳过: %s", exc)

    ctx.px_challenge_meta = challenge_meta
