"""Vercel serverless entry point for the SuperBryn call quality dashboard.

All routes (the dashboard page and every /api/* endpoint) are rewritten to this one
function by vercel.json. Aggregate data (call counts, verdicts, issue codes) comes from
the compact data/snapshot.json committed to the repo - see build_snapshot.py for how
that's built and refreshed. Per-call transcript/audio is fetched live from the SuperBryn
API on each request - nothing about an individual call's content is stored here, only
the aggregate. SuperBryn issues one API key per agent line, so the key used is picked
per call via SUPERBRYN_API_KEY_<AGENT_LABEL> Vercel environment variables (see
api_key_for() below), falling back to a single SUPERBRYN_API_KEY if set.
"""
import http.client
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler
from zoneinfo import ZoneInfo

API_HOST = "api.superbryn.com"
BASE_PATH = "/public-api/v1"
TZ = ZoneInfo("Asia/Kolkata")
REQUEST_TIMEOUT_SECONDS = 10

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# SuperBryn issues a separate API key per agent line (sales_inbound, sales_outbound, ...) -
# the fetch scripts already take a --api-key per --label for this reason. Set
# SUPERBRYN_API_KEY_<AGENT_LABEL upper-cased> per agent as a Vercel env var; SUPERBRYN_API_KEY
# (no suffix) is an optional fallback for a single-key setup or an unrecognized agent_label.
DEFAULT_API_KEY = os.environ.get("SUPERBRYN_API_KEY")


def api_key_for(agent_label):
    if agent_label:
        scoped = os.environ.get(f"SUPERBRYN_API_KEY_{agent_label.upper()}")
        if scoped:
            return scoped
    return DEFAULT_API_KEY


CALL_AUDIO_RE = re.compile(r"^/api/calls/([0-9a-fA-F-]+)/audio$")
CALL_DETAIL_RE = re.compile(r"^/api/calls/([0-9a-fA-F-]+)$")


def _load_snapshot():
    path = os.path.join(ROOT, "data", "snapshot.json")
    with open(path) as f:
        payload = json.load(f)
    return payload["data"], payload["issue_defs"]


def _load_themes():
    path = os.path.join(ROOT, "data", "root_cause_themes.json")
    if not os.path.exists(path):
        return []
    with open(path) as f:
        return json.load(f)["themes"]


# Module-level: runs once per cold start, reused across invocations on a warm instance.
CALLS, ISSUE_DEFS = _load_snapshot()
CALLS_BY_ID = {r["id"]: r for r in CALLS}
THEMES = _load_themes()


def is_red_alert(rec):
    return rec.get("audit_verdict") in ("TP", "FN")


def call_ts(rec):
    return datetime.fromisoformat(rec["started_at"].replace("Z", "+00:00")).astimezone(TZ)


def day_key(rec):
    return call_ts(rec).date().isoformat()


def agent_key(rec):
    return rec.get("agent_label") or "unknown"


def filtered_calls(days_param, agent_param=None):
    all_calls = CALLS
    if agent_param:
        all_calls = [r for r in all_calls if agent_key(r) == agent_param]
    if not days_param or not all_calls:
        return all_calls
    try:
        n = int(days_param)
    except ValueError:
        return all_calls
    max_day = max(call_ts(r).date() for r in all_calls)
    cutoff = max_day - timedelta(days=n - 1)
    return [r for r in all_calls if call_ts(r).date() >= cutoff]


def compute_verdict_counts(calls):
    c = Counter(r.get("audit_verdict") for r in calls)
    return {"tp": c["TP"], "fp": c["FP"], "fn": c["FN"], "tn": c["TN"]}


def compute_summary(calls):
    total = len(calls)
    v = compute_verdict_counts(calls)
    audited = v["tp"] + v["fp"] + v["fn"] + v["tn"]
    red = v["tp"] + v["fn"]
    esc = sum(1 for r in calls if "requested_escalation" in r["issue_codes"])
    total_cost = sum(r.get("cost_usd") or 0 for r in calls)
    return {
        "total_calls": total,
        "audited_calls": audited,
        "audited_rate": round(audited / total * 100, 1) if total else 0,
        "red_alerts": red,
        "red_rate_of_audited": round(red / audited * 100, 1) if audited else 0,
        "red_rate_of_total": round(red / total * 100, 1) if total else 0,
        "clean_rate": round((total - red) / total * 100, 1) if total else 0,
        "tp": v["tp"], "fp": v["fp"], "fn": v["fn"], "tn": v["tn"],
        "precision": round(v["tp"] / (v["tp"] + v["fp"]) * 100, 1) if (v["tp"] + v["fp"]) else None,
        "recall": round(v["tp"] / (v["tp"] + v["fn"]) * 100, 1) if (v["tp"] + v["fn"]) else None,
        "escalation_calls": esc,
        "escalation_rate": round(esc / total * 100, 1) if total else 0,
        "total_cost_usd": round(total_cost, 2),
        "avg_cost_usd": round(total_cost / total, 4) if total else 0,
    }


def compute_days(calls):
    buckets = defaultdict(lambda: {"total": 0, "audited": 0, "red": 0})
    for r in calls:
        b = buckets[day_key(r)]
        b["total"] += 1
        if r.get("audit_verdict") is not None:
            b["audited"] += 1
        if is_red_alert(r):
            b["red"] += 1
    days = []
    for day in sorted(buckets):
        b = buckets[day]
        days.append({
            "day": day,
            "total": b["total"],
            "audited": b["audited"],
            "red": b["red"],
            "red_rate_of_audited": round(b["red"] / b["audited"] * 100, 1) if b["audited"] else 0,
            "red_rate_of_total": round(b["red"] / b["total"] * 100, 1) if b["total"] else 0,
        })
    return days


def compute_issues(calls):
    considered = sum(1 for r in calls if r["issue_codes"] or r.get("audit_verdict") is not None)
    groups = defaultdict(lambda: {"call_ids": []})
    for r in calls:
        for code in r["issue_codes"]:
            groups[code]["call_ids"].append(r["id"])
    issues = []
    for code, g in groups.items():
        d = ISSUE_DEFS.get(code, {})
        count = len(g["call_ids"])
        issues.append({
            "code": code,
            "title": d.get("title", code),
            "lane": d.get("lane"),
            "severity": d.get("severity"),
            "call_ids": g["call_ids"],
            "count": count,
            "pct_of_audited": round(count / considered * 100, 1) if considered else 0,
        })
    issues.sort(key=lambda g: -g["count"])
    return issues


def compute_meta():
    if not CALLS:
        return {"total_calls": 0, "min_day": None, "max_day": None, "days_available": 0, "agents": []}
    days = sorted({day_key(r) for r in CALLS})
    agents = sorted({agent_key(r) for r in CALLS})
    return {
        "total_calls": len(CALLS),
        "min_day": days[0],
        "max_day": days[-1],
        "days_available": len(days),
        "agents": agents,
    }


def call_brief(rec):
    return {
        "id": rec["id"],
        "started_at": rec["started_at"],
        "duration_seconds": rec.get("duration_seconds"),
        "agent": agent_key(rec),
        "cost_usd": rec.get("cost_usd"),
        "has_recording": rec.get("has_recording"),
    }


def fetch_call_detail(call_id, api_key):
    if not api_key:
        return 500, json.dumps({"error": {"message": "Server missing a SuperBryn API key for this call's agent", "code": "no_api_key"}}).encode()
    req = urllib.request.Request(
        f"https://{API_HOST}{BASE_PATH}/observability/calls/{call_id}",
        headers={"X-API-Key": api_key},
    )
    try:
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT_SECONDS) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()
    except urllib.error.URLError as e:
        return 504, json.dumps({"error": {"message": f"Upstream request failed: {e.reason}", "code": "upstream_unreachable"}}).encode()


def fetch_audio_redirect(call_id, api_key):
    if not api_key:
        return 500, None
    conn = http.client.HTTPSConnection(API_HOST, timeout=REQUEST_TIMEOUT_SECONDS)
    try:
        conn.request("GET", f"{BASE_PATH}/observability/calls/{call_id}/audio", headers={"X-API-Key": api_key})
        resp = conn.getresponse()
        location = resp.getheader("Location")
        status = resp.status
        resp.read()
        return status, location
    except OSError:
        return 504, None
    finally:
        conn.close()


class handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def _send_json(self, obj, status=200):
        body = json.dumps(obj).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_bytes(self, body, status, content_type):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _serve_dashboard(self):
        path = os.path.join(ROOT, "dashboard.html")
        with open(path, "rb") as f:
            self._send_bytes(f.read(), 200, "text/html")

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)
        days_param = qs.get("days", [None])[0]
        agent_param = qs.get("agent", [None])[0]

        if path in ("/", "/index.html", "/api/index"):
            self._serve_dashboard()
            return
        if path == "/api/meta":
            self._send_json(compute_meta())
            return
        if path == "/api/summary":
            self._send_json(compute_summary(filtered_calls(days_param, agent_param)))
            return
        if path == "/api/days":
            self._send_json(compute_days(filtered_calls(days_param, agent_param)))
            return
        if path == "/api/issues":
            self._send_json(compute_issues(filtered_calls(days_param, agent_param)))
            return
        if path == "/api/themes":
            self._send_json(THEMES)
            return
        if path == "/api/calls" and "ids" in qs:
            ids = qs["ids"][0].split(",")
            briefs = [call_brief(CALLS_BY_ID[i]) for i in ids if i in CALLS_BY_ID]
            self._send_json(briefs)
            return

        m = CALL_AUDIO_RE.match(path)
        if m:
            call_id = m.group(1)
            rec = CALLS_BY_ID.get(call_id)
            key = api_key_for(rec.get("agent_label") if rec else None)
            status, location = fetch_audio_redirect(call_id, key)
            if status == 302 and location:
                self.send_response(302)
                self.send_header("Location", location)
                self.end_headers()
            else:
                self._send_json({"error": "audio not available"}, status=404)
            return

        m = CALL_DETAIL_RE.match(path)
        if m:
            call_id = m.group(1)
            rec = CALLS_BY_ID.get(call_id)
            key = api_key_for(rec.get("agent_label") if rec else None)
            status, body = fetch_call_detail(call_id, key)
            self._send_bytes(body, status, "application/json")
            return

        self._send_json({"error": "not found"}, status=404)
