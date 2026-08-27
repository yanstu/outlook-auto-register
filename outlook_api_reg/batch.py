"""批量并发注册。

设计要点：
- 线程池并发（注册瓶颈在网络 IO + captcha.run 等待，线程池即可，避免多进程序列化开销）。
- 每个 worker 内部各自 new 一个 OutlookHttpSession（见 register_one → _register_attempt），
  代理池由 register_one 自行轮换；收码池由 proof_pool.iter_accounts 进程内游标轮换（线程安全）。
- 进度既支持回调 on_progress(event: dict)，也提供生成器 register_batch_iter 便于网页/CLI 消费。
- 线程安全：结果与计数用锁保护；保存账号各自写文件（save_account 内 append 有 GIL 足够）。

网页可直接 import：
    from outlook_api_reg.batch import register_batch, register_batch_iter
"""
from __future__ import annotations

import logging
import os
import queue
import random
import threading
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, as_completed, wait
from typing import Callable, Iterator, Optional

from .models import RegisterResult
from .proxy_utils import expand_proxy_unique, has_sid_template
from .register import register_one, save_account

logger = logging.getLogger(__name__)

ProgressCb = Callable[[dict], None]

# 注册启动错峰（秒）默认区间：相邻账号的注册在此范围内随机间隔启动，
# 避免同一时刻爆发式注册（可被显式参数或环境变量 OUTLOOK_REG_JITTER_MIN/MAX 覆盖）。
_DEFAULT_JITTER_MIN = 3.0
_DEFAULT_JITTER_MAX = 8.0


def _resolve_jitter(jitter_min: Optional[float], jitter_max: Optional[float]) -> tuple[float, float]:
    """解析注册间隔抖动：显式参数 > 环境变量 OUTLOOK_REG_JITTER_MIN/MAX > 内置默认。"""
    def _env(name: str, default: float) -> float:
        try:
            return float((os.environ.get(name) or "").strip() or default)
        except (TypeError, ValueError):
            return default

    lo = float(jitter_min) if jitter_min is not None else _env("OUTLOOK_REG_JITTER_MIN", _DEFAULT_JITTER_MIN)
    hi = float(jitter_max) if jitter_max is not None else _env("OUTLOOK_REG_JITTER_MAX", _DEFAULT_JITTER_MAX)
    lo = max(0.0, lo)
    hi = max(lo, hi)
    return lo, hi


def _plan_proxies(proxy: Optional[str], count: int) -> list[Optional[str]]:
    """为每个账号规划一条代理会话：含 `{sid}` → 每号唯一 sid（不同出口 IP）。

    无代理时全 None（直连）；无 `{sid}` 时全批同串（同 IP，调用方另行警告）。
    """
    count = max(1, count)
    if not (proxy and proxy.strip()):
        return [None] * count
    plan = list(expand_proxy_unique(proxy, count))
    if len(plan) < count:  # 兜底：异常输入下补齐，避免 _run_one 索引越界
        plan += [plan[-1] if plan else None] * (count - len(plan))
    return plan


def _make_launch_gate(jitter_min: float, jitter_max: float) -> Callable[[], None]:
    """返回一个 gate()：让相邻账号的注册启动错峰 random.uniform(min,max) 秒。

    线程池内每个 worker 调用一次 gate()，按到达顺序领取递增的启动时隙，
    使得即便并发>1，各账号也不会在同一瞬间同时向微软发起注册。
    """
    lock = threading.Lock()
    next_at = [0.0]

    def gate() -> None:
        if jitter_max <= 0:
            return
        with lock:
            now = time.monotonic()
            start_at = max(now, next_at[0])
            next_at[0] = start_at + random.uniform(jitter_min, jitter_max)
        delay = start_at - time.monotonic()
        if delay > 0:
            time.sleep(delay)

    return gate


def register_batch_iter(
    count: int,
    *,
    concurrency: int = 2,
    email_prefix: Optional[str] = None,
    email_domain: str = "@outlook.com",
    country: str = "SG",
    proxy: Optional[str] = None,
    proxy_plan: Optional[list[Optional[str]]] = None,
    proxy_templates: Optional[list[Optional[str]]] = None,
    px_mode: str = "solver",
    skip_post_login: bool = False,
    fetch_mail_token: bool = True,
    output_dir: Optional[str] = "accounts",
    batch_id: str = "",
    batch_no: Optional[int] = None,
    batch_label: str = "",
    jitter_min: Optional[float] = None,
    jitter_max: Optional[float] = None,
) -> Iterator[dict]:
    """并发注册 count 个账号，yield 进度事件（便于网页 SSE / CLI 消费）。

    防封内建：
      - 一号一 IP：含 `{sid}` 的代理模板为每个账号展开唯一 sticky 会话（出口 IP 互不相同）；
        缺 `{sid}` 时日志警告「全批共用同一 IP，封号风险高」。
      - 启动错峰：相邻账号注册按 jitter_min~jitter_max 秒随机间隔启动，避免爆发式注册。

    事件 event 形如：
      {"type":"start","total":N,"concurrency":C,"proxy_unique":bool,"jitter":[lo,hi]}
      {"type":"account_start","index":i}  # 错峰 gate 通过后、register_one 开始前
      {"type":"result","index":i,"success":bool,"email":..,"combo":..,"combo_dual":..,
       "refresh_token_present":bool,"login_token_present":bool,"elapsed":s,"timings":{..},"error":..}
      {"type":"done","total":N,"ok":k,"failed":m,"elapsed":s,"avg_per_account":s,
       "avg_stage_timings":{..}}
    """
    count = max(1, int(count))
    concurrency = max(1, min(int(concurrency), count))
    jmin, jmax = _resolve_jitter(jitter_min, jitter_max)
    if proxy_plan is not None:
        proxy_plan = list(proxy_plan)
        if len(proxy_plan) < count:
            filler = proxy_plan[-1] if proxy_plan else None
            proxy_plan += [filler] * (count - len(proxy_plan))
        proxy_unique = any(has_sid_template(p) for p in proxy_plan if p)
    else:
        proxy_plan = _plan_proxies(proxy, count)
        proxy_unique = bool(proxy and has_sid_template(proxy))

    if proxy_plan is not None:
        using_pool_plan = True
        if proxy and not proxy_unique and not any(has_sid_template(x) for x in proxy_plan if x):
            logger.info("代理池批量：%d 条代理已按池策略分配。", count)
    elif proxy and not has_sid_template(proxy):
        logger.warning(
            "当前代理会使本批 %d 个账号共用同一出口，容易被一起拦截。请改用独立出口。",
            count,
        )
    elif proxy_unique:
        logger.info("独立出口：已为 %d 个账号各分配独立出口。", count)
    if jmax > 0:
        logger.info("启动间隔：相邻账号注册随机间隔 %.1f–%.1f 秒启动。", jmin, jmax)

    yield {
        "type": "start",
        "total": count,
        "concurrency": concurrency,
        "proxy_unique": proxy_unique,
        "proxy_has_sid": proxy_unique or bool(proxy and has_sid_template(proxy)),
        "jitter": [jmin, jmax],
    }

    gate = _make_launch_gate(jmin, jmax)
    t_batch = time.perf_counter()
    lock = threading.Lock()
    progress_q: queue.Queue[dict] = queue.Queue()
    ok = 0
    dual_req = 0
    dual_ok = 0
    stage_sum: dict[str, float] = {}
    stage_n: dict[str, int] = {}

    def _run_one(idx: int) -> tuple[int, RegisterResult, float]:
        gate()  # 启动错峰：领取本账号的启动时隙，避免同一时刻爆发式注册
        progress_q.put({"type": "account_start", "index": idx})
        t0 = time.perf_counter()
        res = register_one(
            email_prefix=email_prefix,
            email_domain=email_domain,
            country=country,
            proxy=proxy_plan[idx],
            proxy_template=(
                proxy_templates[idx]
                if proxy_templates is not None and idx < len(proxy_templates)
                else None
            ),
            px_mode=px_mode,
            skip_post_login=skip_post_login,
            fetch_mail_token=fetch_mail_token,
        )
        if res.success and output_dir:
            try:
                save_account(
                    res, output_dir,
                    batch_id=batch_id, batch_no=batch_no, batch_label=batch_label,
                )
            except Exception as exc:  # noqa: BLE001
                logger.warning("保存账号失败: %s", exc)
        return idx, res, round(time.perf_counter() - t0, 2)

    def _drain_progress() -> Iterator[dict]:
        while True:
            try:
                yield progress_q.get_nowait()
            except queue.Empty:
                break

    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        pending = {pool.submit(_run_one, i): i for i in range(count)}
        while pending:
            yield from _drain_progress()
            done, _ = wait(pending.keys(), timeout=0.35, return_when=FIRST_COMPLETED)
            if not done:
                continue
            for fut in done:
                pending.pop(fut, None)
                idx, res, elapsed = fut.result()
                timings = (res.extra or {}).get("timings", {}) if res.extra else {}
                with lock:
                    if res.success:
                        ok += 1
                    if (res.extra or {}).get("dual_requested"):
                        dual_req += 1
                        if (res.extra or {}).get("dual_ok"):
                            dual_ok += 1
                    for k, v in (timings or {}).items():
                        stage_sum[k] = stage_sum.get(k, 0.0) + float(v)
                        stage_n[k] = stage_n.get(k, 0) + 1
                yield {
                    "type": "result",
                    "index": idx,
                    "success": bool(res.success),
                    "email": res.email,
                    "combo": res.to_combo() if res.success else "",
                    "combo_dual": res.to_combo(dual=True) if (res.success and res.login_refresh_token) else "",
                    "combo_recovery": res.to_combo(recovery=True) if (res.success and res.recovery_email) else "",
                    "recovery_email": res.recovery_email if res.success else "",
                    "recovery_password": res.recovery_password if res.success else "",
                    "refresh_token_present": bool(res.refresh_token),
                    "login_token_present": bool(res.login_refresh_token),
                    "has_recovery": bool(res.recovery_email and res.recovery_password),
                    "dual_requested": bool((res.extra or {}).get("dual_requested")),
                    "dual_ok": bool((res.extra or {}).get("dual_ok")),
                    "login_status": (res.extra or {}).get("login_status", ""),
                    "login_fail_reason": (res.extra or {}).get("login_fail_reason", ""),
                    "elapsed": elapsed,
                    "timings": timings,
                    "error": res.error,
                }
        yield from _drain_progress()

    total_elapsed = round(time.perf_counter() - t_batch, 2)
    avg_stage = {k: round(stage_sum[k] / stage_n[k], 2) for k in stage_sum if stage_n.get(k)}
    if ok:
        logger.info(
            "本批 %d 个新号已入库，请勿立刻测活。建议静置数小时后再首次校验。",
            ok,
        )
    yield {
        "type": "done",
        "total": count,
        "ok": ok,
        "failed": count - ok,
        "dual_requested": dual_req,
        "dual_ok": dual_ok,
        "elapsed": total_elapsed,
        "avg_per_account": round(total_elapsed / count, 2),
        "avg_stage_timings": avg_stage,
    }


def register_batch(
    count: int,
    *,
    concurrency: int = 2,
    on_progress: Optional[ProgressCb] = None,
    **kwargs,
) -> list[RegisterResult]:
    """并发注册 count 个账号，返回 RegisterResult 列表。

    on_progress 会在每个事件（start/result/done）时被调用一次（可选）。
    其余参数透传给 register_batch_iter（email_domain/country/proxy/px_mode/
    skip_post_login/fetch_mail_token/output_dir/jitter_min/jitter_max 等）。
    防封同 register_batch_iter：一号一 IP（唯一 {sid}）+ 启动错峰。
    """
    results: dict[int, RegisterResult] = {}
    # 为了返回 RegisterResult，需要在迭代里同时保留对象；这里重跑一遍逻辑不划算，
    # 改为在 iter 内不返回对象——故这里直接用一个内部收集器。
    # 简化实现：复用 iter，但结果对象通过 result 事件的 combo 无法还原完整对象，
    # 因此单独跑一遍带对象收集的路径。
    count = max(1, int(count))
    concurrency = max(1, min(int(concurrency), count))
    proxy = kwargs.get("proxy")
    jmin, jmax = _resolve_jitter(kwargs.get("jitter_min"), kwargs.get("jitter_max"))
    proxy_plan = _plan_proxies(proxy, count)
    if proxy and not has_sid_template(proxy):
        logger.warning(
            "当前代理会使本批 %d 个账号共用同一出口，容易被一起拦截。", count
        )
    gate = _make_launch_gate(jmin, jmax)
    if on_progress:
        on_progress({"type": "start", "total": count, "concurrency": concurrency})

    t_batch = time.perf_counter()
    lock = threading.Lock()
    output_dir = kwargs.get("output_dir", "accounts")

    def _run_one(idx: int) -> tuple[int, RegisterResult, float]:
        gate()
        t0 = time.perf_counter()
        res = register_one(
            email_prefix=kwargs.get("email_prefix"),
            email_domain=kwargs.get("email_domain", "@outlook.com"),
            country=kwargs.get("country", "SG"),
            proxy=proxy_plan[idx],
            px_mode=kwargs.get("px_mode", "solver"),
            skip_post_login=kwargs.get("skip_post_login", False),
            fetch_mail_token=kwargs.get("fetch_mail_token", True),
        )
        if res.success and output_dir:
            try:
                save_account(
                    res, output_dir,
                    batch_id=str(kwargs.get("batch_id") or ""),
                    batch_no=kwargs.get("batch_no"),
                    batch_label=str(kwargs.get("batch_label") or ""),
                )
            except Exception as exc:  # noqa: BLE001
                logger.warning("保存账号失败: %s", exc)
        return idx, res, round(time.perf_counter() - t0, 2)

    ok = 0
    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futs = [pool.submit(_run_one, i) for i in range(count)]
        for fut in as_completed(futs):
            idx, res, elapsed = fut.result()
            with lock:
                results[idx] = res
                if res.success:
                    ok += 1
            if on_progress:
                on_progress({
                    "type": "result", "index": idx, "success": bool(res.success),
                    "email": res.email, "combo": res.to_combo() if res.success else "",
                    "combo_dual": res.to_combo(dual=True) if (res.success and res.login_refresh_token) else "",
                    "elapsed": elapsed, "error": res.error,
                })

    if on_progress:
        total_elapsed = round(time.perf_counter() - t_batch, 2)
        on_progress({
            "type": "done", "total": count, "ok": ok, "failed": count - ok,
            "elapsed": total_elapsed, "avg_per_account": round(total_elapsed / count, 2),
        })
    return [results[i] for i in sorted(results)]
