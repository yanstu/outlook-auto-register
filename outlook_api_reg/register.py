from __future__ import annotations

import json
import logging
import os
import random
import secrets
import string
import time
from pathlib import Path
from typing import Optional

import requests
from dotenv import load_dotenv

from .api import check_available_signin_name, create_account
from .bootstrap import bootstrap_session, preload_perimeterx
from .constants import DUAL_TOKEN, LOGIN_CLIENT_ID, MAIL_CLIENT_ID, OUTLOOK_EMAIL_DOMAINS, locale_for_country
from .http_session import OutlookHttpSession
from .account_persist import enrich_register_result
from .account_store import save_register_result as save_account
from .models import AccountInfo, RegisterResult, SignupSession
from .post_register import complete_oauth_after_signup
from .risk import RegisterRetryable, solve_risk_challenge

from .proxy_utils import expand_proxy_template, expand_proxy_unique, has_sid_template, parse_proxy_pool, preflight_proxy

logger = logging.getLogger(__name__)
load_dotenv()

_TRANSIENT_ERRORS = (
    requests.exceptions.ProxyError,
    requests.exceptions.SSLError,
    requests.exceptions.ConnectionError,
    requests.exceptions.Timeout,
)


def _is_transient(exc: Exception) -> bool:
    return isinstance(exc, _TRANSIENT_ERRORS)


def _register_attempt(
    *,
    email_prefix: Optional[str],
    email_domain: str,
    country: str,
    proxy: Optional[str],
    px_mode: str,
    skip_post_login: bool,
    fetch_mail_token: bool,
) -> RegisterResult:
    http = OutlookHttpSession(proxy=proxy)

    if email_domain not in OUTLOOK_EMAIL_DOMAINS:
        email_domain = "@outlook.com"

    timings: dict[str, float] = {}

    def _stage(name: str, fn):
        t0 = time.perf_counter()
        try:
            return fn()
        finally:
            timings[name] = round(time.perf_counter() - t0, 2)

    mkt, lc = locale_for_country(country)
    ctx = _stage("bootstrap", lambda: bootstrap_session(http, mkt=mkt, lc=lc))
    _stage("px_preload", lambda: preload_perimeterx(http, ctx))

    prefix = email_prefix or _random_email_prefix()
    email, check_map = _stage(
        "pick_email", lambda: _pick_available_email(http, ctx, prefix, domain=email_domain)
    )
    logger.info("选用邮箱: %s", email)

    first, last = _random_name()
    password = _random_password()
    account = AccountInfo(
        email=email,
        password=password,
        first_name=first,
        last_name=last,
        country=country,
        birth_date=_random_birthday(),
    )

    _stage("risk_px", lambda: solve_risk_challenge(http, ctx, account, mode=px_mode, proxy=proxy))

    create_data = _stage("create_account", lambda: create_account(
        http, ctx, account,
        check_avail_map=check_map,
        member_name_change_count=len(check_map),
        member_name_available_count=1,
        member_name_unavailable_count=len(check_map) - 1,
    ))

    redirect_url = create_data.get("redirectUrl", "")
    slt = create_data.get("slt", "")
    signin_name = create_data.get("signinName", email)

    post_info = {}
    if not skip_post_login and redirect_url and slt:
        post_info = _stage("oauth_mail", lambda: complete_oauth_after_signup(
            http, ctx, redirect_url, slt,
            email=signin_name,
            password=password,
            proxy=proxy,
            fetch_mail_token=fetch_mail_token,
            country=country,
        ))

    logger.info("各阶段耗时：%s", timings)

    # 产出前自检：确认这枚令牌真能读信（只授 IMAP/POP/SMTP 的新号会 401 → 废号）。
    # 用与注册同一条代理，避免读信 IP 与建号 IP 不一致。
    mail_rt = post_info.get("mail_refresh_token", "")
    mail_check: dict = {}
    if mail_rt and fetch_mail_token:
        from .graph_mail import verify_mail_readable
        from .proxy_utils import parse_proxy

        _cfg = parse_proxy(proxy) if proxy else None
        mail_check = verify_mail_readable(
            mail_rt, client_id=MAIL_CLIENT_ID, proxy_url=(_cfg.url if _cfg else ""),
        )
        if mail_check.get("readable"):
            logger.info(
                "✅ 产出自检：令牌可读信 name=%s",
                mail_check.get("display_name"),
            )
        else:
            logger.warning(
                "⚠️ 产出自检：令牌读信失败（%s %s）——该号可能不可用",
                mail_check.get("status") or mail_check.get("reason"),
                mail_check.get("detail", ""),
            )

    login_rt = post_info.get("login_refresh_token", "")
    dual_summary = {}
    if DUAL_TOKEN:
        dual_summary = {
            "dual_requested": True,
            "dual_ok": bool(login_rt),
            "login_status": post_info.get("login_status", "" if login_rt else "missing"),
            "login_fail_reason": post_info.get("login_fail_reason", ""),
        }
    return enrich_register_result(RegisterResult(
        success=True,
        email=signin_name,
        password=password,
        redirect_url=redirect_url,
        slt=slt,
        refresh_token=post_info.get("mail_refresh_token", ""),
        client_id=MAIL_CLIENT_ID,
        login_client_id=LOGIN_CLIENT_ID if login_rt else "",
        login_refresh_token=login_rt,
        recovery_email=post_info.get("recovery_email", ""),
        recovery_password=post_info.get("recovery_password", ""),
        extra={
            "post_login": post_info,
            "uaid": ctx.uaid,
            "timings": timings,
            "proofs_method": post_info.get("proofs_method", ""),
            "proofs_satisfied": post_info.get("proofs_satisfied", ""),
            "mail_readable": mail_check.get("readable"),
            "mail_check": mail_check,
            **dual_summary,
        },
    ))


def _random_email_prefix(length: int = 0) -> str:
    # 对齐卖家风格：纯小写字母、无数字、10-12 位（带数字/年份最像脚本，去掉）。
    n = length or random.randint(10, 12)
    return "".join(random.choice(string.ascii_lowercase) for _ in range(n))


def _random_password(length: int = 0) -> str:
    # 对齐卖家风格：小写字母+数字、11-14 位、无大写无符号（满足 MSA 两类字符要求）。
    n = length or random.randint(11, 14)
    chars = string.ascii_lowercase + string.digits
    while True:
        pwd = "".join(secrets.choice(chars) for _ in range(n))
        if any(c.islower() for c in pwd) and any(c.isdigit() for c in pwd):
            return pwd


# 英美名大词表（对齐卖家：119 个不重复的 First Last，避免 lx/Ren + 亚洲姓小词表一眼农场）。
_FIRST_NAMES = (
    "James", "John", "Robert", "Michael", "William", "David", "Richard", "Joseph",
    "Thomas", "Charles", "Christopher", "Daniel", "Matthew", "Anthony", "Mark",
    "Donald", "Steven", "Paul", "Andrew", "Joshua", "Kenneth", "Kevin", "Brian",
    "George", "Timothy", "Ronald", "Edward", "Jason", "Jeffrey", "Ryan", "Jacob",
    "Gary", "Nicholas", "Eric", "Jonathan", "Stephen", "Larry", "Justin", "Scott",
    "Brandon", "Benjamin", "Samuel", "Gregory", "Alexander", "Patrick", "Jack",
    "Dennis", "Jerry", "Tyler", "Aaron", "Jose", "Nathan", "Adam", "Henry", "Peter",
    "Mary", "Patricia", "Jennifer", "Linda", "Elizabeth", "Barbara", "Susan",
    "Jessica", "Sarah", "Karen", "Nancy", "Lisa", "Betty", "Margaret", "Sandra",
    "Ashley", "Kimberly", "Emily", "Donna", "Michelle", "Carol", "Amanda", "Melissa",
    "Deborah", "Stephanie", "Rebecca", "Laura", "Sharon", "Cynthia", "Kathleen",
    "Amy", "Angela", "Shirley", "Anna", "Brenda", "Pamela", "Nicole", "Samantha",
    "Katherine", "Christine", "Rachel", "Courtney", "Danielle", "Cheyenne", "Veronica",
)
_LAST_NAMES = (
    "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis",
    "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson",
    "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson",
    "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson", "Walker",
    "Young", "Allen", "King", "Wright", "Scott", "Torres", "Nguyen", "Hill", "Flores",
    "Green", "Adams", "Nelson", "Baker", "Hall", "Rivera", "Campbell", "Mitchell",
    "Carter", "Roberts", "Gomez", "Phillips", "Evans", "Turner", "Diaz", "Parker",
    "Cruz", "Edwards", "Collins", "Reyes", "Stewart", "Morris", "Morales", "Murphy",
    "Cook", "Rogers", "Gutierrez", "Ortiz", "Morgan", "Cooper", "Peterson", "Bailey",
    "Reed", "Kelly", "Howard", "Cox", "Ward", "Richardson", "Watson", "Brooks",
    "Wood", "Bennett", "Gray", "Fisher", "Byrd", "Pace", "Flynn", "Odonnell",
    "Bowen", "Fischer", "Decker", "Ball", "Mcgrath", "Lang", "Padilla", "Frost",
    "Buchanan", "Estrada", "Spencer", "Sellers", "Owens", "Hart",
)


def _random_name() -> tuple[str, str]:
    return random.choice(_FIRST_NAMES), random.choice(_LAST_NAMES)


def _random_birthday() -> str:
    year = random.randint(1985, 2000)
    month = random.randint(1, 12)
    day = random.randint(1, 28)
    return f"{day:02d}:{month:02d}:{year}"


def _pick_available_email(
    http: OutlookHttpSession,
    ctx: SignupSession,
    prefix: str,
    *,
    domain: str = "@outlook.com",
    max_tries: int = 6,
) -> tuple[str, list[str]]:
    # 对齐卖家：只用纯小写字母前缀，不再退化成 prefix+数字 / prefix+年份。
    # 10-12 位随机字母几乎不会撞名；撞了就换一个新的纯字母前缀，而不是加数字。
    check_map: list[str] = []
    tried: set[str] = set()

    def _next_prefix(first: bool) -> str:
        return prefix if first else _random_email_prefix()

    for i in range(max_tries):
        p = _next_prefix(i == 0)
        if p in tried:
            continue
        tried.add(p)
        email = f"{p}{domain}"
        data = check_available_signin_name(http, ctx, email)
        available = data.get("isAvailable", False)
        check_map.append(f"{email}:{str(not available).lower()}")
        if available:
            return email, check_map
        # 微软给的建议里挑「纯小写字母」的用，跳过带数字/大写的（保持卖家风格）。
        for sug in data.get("suggestions", [])[:5]:
            local = sug.split("@")[0]
            if not (local.isalpha() and local.islower()):
                continue
            data2 = check_available_signin_name(http, ctx, sug)
            check_map.append(f"{sug}:{str(not data2.get('isAvailable', False)).lower()}")
            if data2.get("isAvailable"):
                return sug, check_map

    raise RuntimeError(f"无法找到可用邮箱（已试 {len(tried)} 个纯字母前缀）")


def register_one(
    *,
    email_prefix: Optional[str] = None,
    email_domain: str = "@outlook.com",
    country: str = "US",
    proxy: Optional[str] = None,
    proxy_template: Optional[str] = None,
    px_mode: str = "solver",
    skip_post_login: bool = False,
    fetch_mail_token: bool = False,
) -> RegisterResult:
    explicit_proxy = proxy or os.environ.get("HTTP_PROXY") or None
    retries = max(1, int((os.environ.get("REG_PROXY_RETRIES") or "3").strip() or "3"))

    tmpl = (proxy_template or explicit_proxy or "").strip()
    if has_sid_template(tmpl):
        attempt_proxies = expand_proxy_unique(tmpl, count=retries)
    elif explicit_proxy:
        # 已解析的 sticky 会话：整号重试时仍用同串，但会新建 PX/captcha 会话
        attempt_proxies = [explicit_proxy] * retries
    else:
        pool = parse_proxy_pool(template_count=retries)
        attempt_proxies = pool if pool else [None]

    last_exc: Optional[Exception] = None
    last_err = ""
    for idx, p in enumerate(attempt_proxies, start=1):
        ok, info = preflight_proxy(p)
        if not ok:
            logger.error("代理预检失败，跳过（%s/%s）: %s", idx, len(attempt_proxies), info)
            last_err = info
            continue
        logger.info("代理预检通过（%s/%s）: %s [%s]", idx, len(attempt_proxies), p or "(直连)", info)
        if len(attempt_proxies) > 1:
            logger.info("注册尝试 %s/%s", idx, len(attempt_proxies))
        try:
            return _register_attempt(
                email_prefix=email_prefix,
                email_domain=email_domain,
                country=country,
                proxy=p,
                px_mode=px_mode,
                skip_post_login=skip_post_login,
                fetch_mail_token=fetch_mail_token,
            )
        except Exception as exc:
            last_exc = exc
            remaining = idx < len(attempt_proxies)
            retryable = _is_transient(exc) or isinstance(exc, RegisterRetryable)
            if remaining and retryable:
                logger.warning(
                    "注册失败，换新会话重试（%s/%s）: %s",
                    idx, len(attempt_proxies), exc,
                )
                time.sleep(2.0)
                continue
            logger.exception("注册失败")
            return RegisterResult(success=False, error=str(exc))

    if last_exc:
        return RegisterResult(success=False, error=str(last_exc))
    if last_err:
        return RegisterResult(
            success=False,
            error=(
                f"所有代理均不可用（{len(attempt_proxies)} 次尝试）。\n  {last_err}"
            ),
        )
    return RegisterResult(success=False, error="未知错误")

