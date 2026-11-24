# -*- coding: utf-8 -*-
"""实时数据历史留痕与走势聚合。

每次实时抓取成功后追加一条快照到 ``<data>/realtime_history/<日期>.jsonl``，
按「当天最后一次快照」得到每个账户的当日累计值，再按日 / 周 / 月聚合成折线图数据。

本模块不依赖浏览器内核，可独立单测。
"""
from __future__ import annotations

import json
import os
import re
import calendar
from datetime import date, datetime, time, timedelta
from pathlib import Path

KEEP_DAYS = 180           # 保留多少天的留痕文件
MAX_LINES_PER_DAY = 800   # 单日快照上限（5 分钟一次约 111 条，足够冗余）
DIMS = ("day", "week", "month")
SCOPES = ("recent", "period")
METRICS = ("cost", "gmv")


# ---------------------------------------------------------------- 路径
def history_dir(root) -> Path:
    return Path(root) / "realtime_history"


def accounts_path(root) -> Path:
    return history_dir(root) / "accounts.json"


def _day_file(root, day: str) -> Path:
    return history_dir(root) / f"{day}.jsonl"


# ---------------------------------------------------------------- 写入
def _now() -> datetime:
    return datetime.now()


def _atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def _load_accounts(root) -> dict:
    p = accounts_path(root)
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:  # noqa: BLE001
        return {}


def _merge_accounts(root, rows) -> None:
    """把本次出现的账户名/所属登录账号补进索引（名字变了就更新）"""
    cur = _load_accounts(root)
    changed = False
    for r in rows:
        key = r.get("key")
        if not key:
            continue
        item = {"name": r.get("name") or "", "owner": r.get("owner") or ""}
        if cur.get(key) != item:
            cur[key] = item
            changed = True
    if changed:
        _atomic_write_text(accounts_path(root), json.dumps(cur, ensure_ascii=False, sort_keys=True))


def _trim_day_file(path: Path) -> None:
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except Exception:  # noqa: BLE001
        return
    if len(lines) <= MAX_LINES_PER_DAY:
        return
    keep = lines[-MAX_LINES_PER_DAY:]
    _atomic_write_text(path, "\n".join(keep) + "\n")


def _prune(root, today: date) -> None:
    """删掉超过 KEEP_DAYS 天的留痕文件"""
    d = history_dir(root)
    if not d.is_dir():
        return
    cutoff = today - timedelta(days=KEEP_DAYS)
    for p in d.glob("*.jsonl"):
        m = re.fullmatch(r"(\d{4}-\d{2}-\d{2})\.jsonl", p.name)
        if not m:
            continue
        try:
            day = datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            continue
        if day < cutoff:
            try:
                p.unlink()
            except OSError:
                pass


def snapshot_of(data: dict) -> dict:
    """把一次抓取结果压成一条留痕（只保留画图与对账需要的字段）"""
    rows = []
    for a in data.get("accounts") or []:
        oid = str(a.get("owner_id") or "")
        aid = str(a.get("account_id") or "")
        if not aid:
            continue
        rows.append({
            "key": f"{oid}|{aid}",
            "name": a.get("account_name") or aid,
            "owner": a.get("owner_label") or "",
            "cost": round(float(a.get("cost") or 0), 2),
            "gmv": round(float(a.get("gmv") or 0), 2),
            "order_cnt": int(a.get("order_cnt") or 0),
        })
    return {
        "ts": data.get("fetched_at") or _now().strftime("%Y-%m-%d %H:%M:%S"),
        "date": data.get("date") or _now().strftime("%Y-%m-%d"),
        "elapsed": float(data.get("elapsed") or 0),
        "failures": len(data.get("failures") or []),
        "rows": rows,
    }


def record(root, data: dict) -> bool:
    """追加一次抓取留痕；返回是否写入成功"""
    snap = snapshot_of(data)
    if not snap["rows"]:
        return False
    day = str(snap["date"])[:10]
    try:
        datetime.strptime(day, "%Y-%m-%d")
    except ValueError:
        day = _now().strftime("%Y-%m-%d")

    line = {
        "t": str(snap["ts"])[11:19] or _now().strftime("%H:%M:%S"),
        "e": round(snap["elapsed"], 1),
        "f": snap["failures"],
        "a": [[r["key"], r["cost"], r["gmv"], r["order_cnt"]] for r in snap["rows"]],
    }
    path = _day_file(root, day)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(line, ensure_ascii=False, separators=(",", ":")) + "\n")
    _merge_accounts(root, snap["rows"])
    _trim_day_file(path)
    _prune(root, date.fromisoformat(day))
    return True


# ---------------------------------------------------------------- 读取
def _read_records(path: Path):
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return
    for ln in text.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        try:
            rec = json.loads(ln)
        except Exception:  # noqa: BLE001
            continue
        if isinstance(rec, dict):
            yield rec


def hourly_last(root, day: str) -> dict[int, dict[str, dict]]:
    """{小时: {账户key: {cost,gmv,order_cnt}}}，取该小时最后一次快照（当日累计值）"""
    out: dict[int, dict[str, dict]] = {}
    for rec in _read_records(_day_file(root, day)):
        try:
            hour = int(str(rec.get("t") or "")[:2])
        except ValueError:
            continue
        if not 0 <= hour <= 23:
            continue
        slot = out.setdefault(hour, {})
        for item in rec.get("a") or []:
            if not isinstance(item, list) or len(item) < 4:
                continue
            slot[str(item[0])] = {
                "cost": float(item[1] or 0),
                "gmv": float(item[2] or 0),
                "order_cnt": int(item[3] or 0),
            }
    return out


def daily_last(root) -> dict[str, dict[str, dict]]:
    """{日期: {账户key: {cost,gmv,order_cnt}}}，取当天最后一次出现值"""
    out: dict[str, dict[str, dict]] = {}
    d = history_dir(root)
    if not d.is_dir():
        return out
    for p in sorted(d.glob("*.jsonl")):
        day = p.stem
        per_day = out.setdefault(day, {})
        try:
            lines = p.read_text(encoding="utf-8").splitlines()
        except Exception:  # noqa: BLE001
            continue
        for ln in lines:
            ln = ln.strip()
            if not ln:
                continue
            try:
                rec = json.loads(ln)
            except Exception:  # noqa: BLE001
                continue
            for item in rec.get("a") or []:
                if not isinstance(item, list) or len(item) < 4:
                    continue
                per_day[item[0]] = {
                    "cost": float(item[1] or 0),
                    "gmv": float(item[2] or 0),
                    "order_cnt": int(item[3] or 0),
                }
        if not per_day:
            out.pop(day, None)
    return out

def _week_start(day: date) -> date:
    return day - timedelta(days=day.weekday())


SCOPE_LABELS = {
    ("day", "recent"): "最近24小时",
    ("day", "period"): "本日",
    ("week", "recent"): "最近7天",
    ("week", "period"): "本周",
    ("month", "recent"): "最近31天",
    ("month", "period"): "本月",
}


def _day_buckets(now: datetime, scope: str, anchor: date | None = None):
    """日视图按小时切桶：最近24小时 / 指定某天（本日）00-23 点"""
    today = anchor or now.date()
    if scope == "period":
        return [datetime.combine(today, time(h)) for h in range(24)]
    end = now.replace(minute=0, second=0, microsecond=0)
    return [end - timedelta(hours=23 - i) for i in range(24)]


def _date_buckets(now: datetime, dim: str, scope: str, anchor: date | None = None,
                  anchor_month: str | None = None):
    """周视图 7 天、月视图 28~31 天（最近 / 本期 / 指定）"""
    today = anchor or now.date()
    if dim == "week":
        if scope == "period":
            start = _week_start(today)
            return [start + timedelta(days=i) for i in range(7)]
        return [today - timedelta(days=6 - i) for i in range(7)]
    if anchor_month:
        y, m = int(anchor_month[:4]), int(anchor_month[5:7])
        last = calendar.monthrange(y, m)[1]
        return [date(y, m, d) for d in range(1, last + 1)]
    if scope == "period":
        last = calendar.monthrange(today.year, today.month)[1]
        return [date(today.year, today.month, d) for d in range(1, last + 1)]
    return [today - timedelta(days=30 - i) for i in range(31)]


def series(root, dim: str = "day", metric: str = "cost", scope: str = "recent",
           anchor: str | None = None, now: datetime | None = None) -> dict:
    """折线图数据：每个账号一条线。

    dim=day  按小时（最近24小时 / 本日 / 指定某天）
    dim=week 按天（最近7天 / 本周 / 指定某周，周一为起始）
    dim=month 按天（最近31天 / 本月 / 指定某月）
    anchor 指定区间：日传 YYYY-MM-DD，周传周内任意一天，月传 YYYY-MM
    """
    dim = (dim or "day").lower()
    metric = (metric or "cost").lower()
    scope = (scope or "recent").lower()
    if dim not in DIMS:
        raise ValueError(f"dim 只能是 {'/'.join(DIMS)}")
    if scope not in SCOPES:
        raise ValueError(f"scope 只能是 {'/'.join(SCOPES)}")
    if metric not in METRICS:
        raise ValueError(f"metric 只能是 {'/'.join(METRICS)}")

    now = now or _now()
    accounts = _load_accounts(root)
    bucket_meta: list[dict] = []
    bucket_values: list[dict[str, float]] = []

    anchor_day: date | None = None
    anchor_month: str | None = None
    if anchor:
        anchor = str(anchor).strip()
        if dim == "month" and re.fullmatch(r"\d{4}-\d{2}", anchor):
            anchor_month = anchor
        else:
            try:
                anchor_day = date.fromisoformat(anchor[:10])
            except ValueError:
                raise ValueError("anchor 格式应为 YYYY-MM-DD 或 YYYY-MM") from None

    if dim == "day":
        hours = _day_buckets(now, scope, anchor_day)
        cache: dict[str, dict[int, dict[str, dict]]] = {}
        carry: dict[str, dict[str, float]] = {}   # 当日累计值：无快照的小时沿用上一次
        for dt in hours:
            d = dt.date().isoformat()
            if d not in cache:
                cache[d] = hourly_last(root, d)
            slot = cache[d].get(dt.hour) or {}
            day_carry = carry.setdefault(d, {})
            for k, v in slot.items():
                day_carry[k] = float(v.get(metric) or 0)
            bucket_meta.append({"key": dt.strftime("%Y-%m-%d %H"), "label": dt.strftime("%H:%M")})
            bucket_values.append(dict(day_carry))
    else:
        days = _date_buckets(now, dim, scope, anchor_day, anchor_month)
        daily = daily_last(root)
        for d in days:
            slot = daily.get(d.isoformat()) or {}
            bucket_meta.append({"key": d.isoformat(), "label": d.strftime("%m-%d")})
            bucket_values.append({k: float(v.get(metric) or 0) for k, v in slot.items()})

    all_keys: set[str] = set()
    for vals in bucket_values:
        all_keys.update(vals.keys())

    out_series = []
    for acc_key in all_keys:
        values = [round(vals[acc_key], 2) if acc_key in vals else None for vals in bucket_values]
        meta = accounts.get(acc_key) or {}
        out_series.append({
            "key": acc_key,
            "name": meta.get("name") or acc_key.split("|")[-1],
            "owner": meta.get("owner") or "",
            "values": values,
            "total": round(sum(v for v in values if v), 2),
        })
    out_series.sort(key=lambda s: (-s["total"], s["name"]))

    daily = daily_last(root)
    records = 0
    dd = history_dir(root)
    if dd.is_dir():
        for p in dd.glob("*.jsonl"):
            try:
                records += len(p.read_text(encoding="utf-8").splitlines())
            except OSError:
                pass

    span = bucket_meta[0]["key"] + " ~ " + bucket_meta[-1]["key"] if bucket_meta else ""
    if scope == "recent":
        range_label = SCOPE_LABELS[(dim, "recent")]
    elif anchor and dim == "day":
        range_label = f"{anchor_day.isoformat()} 全天"
    elif anchor and dim == "week":
        start = _week_start(anchor_day)
        range_label = f"{start.isoformat()} ~ {(start + timedelta(days=6)).isoformat()}"
    elif anchor and dim == "month":
        range_label = f"{anchor_month} 全月"
    else:
        range_label = SCOPE_LABELS[(dim, "period")]
    return {
        "ok": True,
        "dim": dim,
        "scope": scope,
        "metric": metric,
        "scope_label": range_label,
        "available": periods(root),
        "buckets": bucket_meta,
        "series": out_series,
        "span": span,
        "days": len(daily),
        "records": records,
        "last_date": max(daily) if daily else "",
        "updated_at": _now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def periods(root) -> dict:
    """有留痕的日 / 周(周一起) / 月清单，供页面「指定区间」下拉使用"""
    daily = daily_last(root)
    days = sorted(daily)
    week_keys, month_keys = set(), set()
    for d in days:
        dt = date.fromisoformat(d)
        week_keys.add(_week_start(dt).isoformat())
        month_keys.add(dt.strftime("%Y-%m"))
    return {
        "days": [{"key": d, "label": d} for d in reversed(days)],
        "weeks": [
            {"key": k, "label": f"{k} ~ {(date.fromisoformat(k) + timedelta(days=6)).strftime('%m-%d')}"}
            for k in sorted(week_keys, reverse=True)
        ],
        "months": [
            {"key": k, "label": f"{k[:4]}年{int(k[5:7])}月"} for k in sorted(month_keys, reverse=True)
        ],
    }
