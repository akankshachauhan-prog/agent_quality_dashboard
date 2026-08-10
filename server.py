#!/usr/bin/env python3
"""Local dashboard server for SuperBryn call quality data.

Serves the dashboard UI and proxies live call-detail / audio requests to the
SuperBryn API so the API key never reaches the browser.

SuperBryn issues one API key per agent line, so set SUPERBRYN_API_KEY_<AGENT_LABEL>
(upper-cased, e.g. SUPERBRYN_API_KEY_SALES_INBOUND) per agent; SUPERBRYN_API_KEY with
no suffix is an optional fallback used when a call's agent has no scoped key set.

Usage:
    export SUPERBRYN_API_KEY_SALES_INBOUND=sbryn_...
    export SUPERBRYN_API_KEY_SALES_OUTBOUND=sbryn_...
    python3 server.py [--port 8000] [--data data/observability_calls_....json]

By default it loads and merges every data/observability_calls_*.json file
present (e.g. separate fetches for different agents/keys), deduped by call id,
then layers in per-call detail (transcript, metric_evaluations, metrics, silence)
from every data/red_alert_transcripts_*.json file for the calls that have it.
"""
import argparse
import glob
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
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from zoneinfo import ZoneInfo

API_HOST = "api.superbryn.com"
BASE_PATH = "/public-api/v1"
TZ = ZoneInfo("Asia/Kolkata")

CALLS = {}
ISSUE_RCA = {}
THEMES = []

# SuperBryn issues a separate API key per agent line (sales_inbound, sales_outbound, ...) -
# fetch_superbryn_calls.py / fetch_red_alert_transcripts.py already take a --api-key per
# --label for this reason. Set SUPERBRYN_API_KEY_<AGENT_LABEL upper-cased> per agent;
# SUPERBRYN_API_KEY (no suffix) is an optional fallback for a single-key setup.
DEFAULT_API_KEY = None


def api_key_for(agent_label):
    if agent_label:
        scoped = os.environ.get(f"SUPERBRYN_API_KEY_{agent_label.upper()}")
        if scoped:
            return scoped
    return DEFAULT_API_KEY


CALL_AUDIO_RE = re.compile(r"^/api/calls/([0-9a-fA-F-]+)/audio$")
CALL_DETAIL_RE = re.compile(r"^/api/calls/([0-9a-fA-F-]+)$")


def load_dataset(path):
    with open(path) as f:
        payload = json.load(f)
    for rec in payload["data"]:
        CALLS[rec["id"]] = rec
    print(f"Loaded {len(payload['data'])} calls from {path} ({len(CALLS)} unique so far)")


def load_call_detail(path):
    """Merge in full call-detail records (transcript, metric_evaluations, metrics, silence,
    audit) fetched separately for red-alert calls by fetch_red_alert_transcripts.py. These
    lack agent_label (the list-view-only field fetch_superbryn_calls.py stamps on), so merge
    onto the existing list-view record rather than replacing it, keeping agent_label intact."""
    with open(path) as f:
        payload = json.load(f)
    n = 0
    for rec in payload["data"]:
        CALLS[rec["id"]] = {**CALLS.get(rec["id"], {}), **rec}
        n += 1
    print(f"Merged detail for {n} calls from {path}")


def call_issues(rec):
    """Per-call flagged issues, bridging two API schema generations: older list-view fetches
    populate analysis_issues directly; newer full-detail records instead carry metric_evaluations
    with one entry per check, of which only result == 'flagged' entries are actual issues."""
    if rec.get("analysis_issues"):
        return rec["analysis_issues"]
    return [e for e in (rec.get("metric_evaluations") or []) if e.get("result") == "flagged"]


def has_call_detail(rec):
    """True if we have per-call issue detail for this call (either schema) - i.e. it's one of
    the audited calls we've fetched full detail for, not just a bare list-view record."""
    return bool(rec.get("analysis_issues") or rec.get("metric_evaluations"))


def call_satisfaction(rec):
    """True/False if the audited call's intent was satisfied, None if we don't have that
    verdict for this call (no full detail fetched, or the field wasn't populated)."""
    ui = ((rec.get("audit") or {}).get("user_intent")) or {}
    s = ui.get("satisfaction")
    if s is None:
        return None
    return s == "true" or s is True


def agent_made_mistake(rec):
    """Whether the agent itself is at fault on this call - any flagged check outside the
    user_behaviour lane (which reflects the customer's behavior, not the agent's)."""
    return any(e.get("lane") != "user_behaviour" for e in call_issues(rec))


def load_issue_rca(path):
    if not os.path.exists(path):
        return
    with open(path) as f:
        entries = json.load(f)
    for entry in entries:
        ISSUE_RCA[entry["code"]] = entry
    print(f"Loaded RCA for {len(entries)} issue buckets from {path}")


def load_root_cause_themes(path):
    """The 9 cross-cutting themes (owner/priority/recommended fix) from
    build_root_cause_sheet.py's THEMES, exported to JSON by that script - same rollup that
    goes into reports/root_cause_analysis.xlsx. Also stamps each ISSUE_RCA root-cause bullet
    with the theme it was assigned to, so the issue drawer can show owner/priority in context."""
    global THEMES
    if not os.path.exists(path):
        return
    with open(path) as f:
        data = json.load(f)
    THEMES = data["themes"]
    theme_by_key = {t["key"]: t for t in THEMES}
    for code, theme_keys in data["issue_theme_by_code"].items():
        bucket = ISSUE_RCA.get(code)
        if not bucket:
            continue
        for rc, theme_key in zip(bucket.get("root_causes") or [], theme_keys):
            t = theme_by_key.get(theme_key)
            rc["theme"] = {"key": t["key"], "name": t["name"], "owner": t["owner"], "priority": t["priority"]} if t else None
    print(f"Loaded {len(THEMES)} root-cause themes from {path}")


def is_red_alert(rec):
    return rec.get("audit_verdict") in ("TP", "FN")


def call_ts(rec):
    return datetime.fromisoformat(rec["started_at"].replace("Z", "+00:00")).astimezone(TZ)


def day_key(rec):
    return call_ts(rec).date().isoformat()


def agent_key(rec):
    """Groups calls by agent. Prefers agent_label (stamped by fetch_superbryn_calls.py
    --label), since the API's own agent_name is often the same generic deployment name
    (e.g. "livekit-agent") across totally different agents/keys."""
    return rec.get("agent_label") or rec.get("agent_name") or "unknown"


def filtered_calls(days_param, agent_param=None):
    """All loaded calls, optionally restricted to the trailing N days (IST) and/or
    a single agent. Day filtering anchors to the most recent call in the dataset
    rather than wall-clock now."""
    all_calls = list(CALLS.values())
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
    red = v["tp"] + v["fn"]  # ground-truth red alerts, only knowable among audited calls
    # Only detectable on calls with per-call issue detail (red-alert calls with merged
    # transcripts) - calls with no analysis_issues/metric_evaluations can't be checked either way.
    esc = sum(1 for r in calls if any(i.get("code") == "requested_escalation" for i in call_issues(r)))
    total_cost = sum(r.get("cost_usd") or 0 for r in calls)
    return {
        "total_calls": total,
        "audited_calls": audited,
        "audited_rate": round(audited / total * 100, 1) if total else 0,
        "red_alerts": red,
        # Ground-truth failure rate within the reviewed sample - the only number backed by a human verdict.
        "red_rate_of_audited": round(red / audited * 100, 1) if audited else 0,
        # Lower bound over the whole population - assumes every never-audited call was clean, which it wasn't.
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
        days.append(
            {
                "day": day,
                "total": b["total"],
                "audited": b["audited"],
                "red": b["red"],
                # Ground-truth rate within the audited sample for that day - use this as the primary series.
                "red_rate_of_audited": round(b["red"] / b["audited"] * 100, 1) if b["audited"] else 0,
                # Lower bound over all calls that day - kept for comparison, not as the headline number.
                "red_rate_of_total": round(b["red"] / b["total"] * 100, 1) if b["total"] else 0,
            }
        )
    return days


def compute_agents(calls):
    buckets = defaultdict(lambda: {"total": 0, "red": 0, "cost": 0.0})
    for r in calls:
        b = buckets[agent_key(r)]
        b["total"] += 1
        if is_red_alert(r):
            b["red"] += 1
        b["cost"] += r.get("cost_usd") or 0
    agents = []
    for name, b in buckets.items():
        agents.append(
            {
                "agent": name,
                "total": b["total"],
                "red": b["red"],
                "red_rate_of_total": round(b["red"] / b["total"] * 100, 1) if b["total"] else 0,
                "total_cost_usd": round(b["cost"], 2),
                "avg_cost_usd": round(b["cost"] / b["total"], 4) if b["total"] else 0,
            }
        )
    return sorted(agents, key=lambda a: -a["total"])


def compute_lanes(calls):
    """Issue-instance counts by lane across every audited call we have detail for (not just
    confirmed red alerts) - a call can be flagged without the customer having noticed."""
    counts = defaultdict(int)
    for r in calls:
        if not has_call_detail(r):
            continue
        for i in call_issues(r):
            counts[i["lane"]] += 1
    return [{"lane": k, "count": v} for k, v in sorted(counts.items(), key=lambda kv: -kv[1])]


def compute_issues(calls):
    considered = sum(1 for r in calls if has_call_detail(r))
    groups = defaultdict(lambda: {"title": None, "lane": None, "code": None, "severity": None, "call_ids": []})
    for r in calls:
        if not has_call_detail(r):
            continue
        for i in call_issues(r):
            g = groups[i["code"]]
            g["title"] = i["title"]
            g["lane"] = i["lane"]
            g["code"] = i["code"]
            g["severity"] = i.get("severity")
            g["call_ids"].append(r["id"])
    issues = list(groups.values())
    for g in issues:
        g["count"] = len(g["call_ids"])
        g["pct_of_audited"] = round(g["count"] / considered * 100, 1) if considered else 0
    issues.sort(key=lambda g: -g["count"])
    return issues


def compute_quadrants(calls):
    """How each audited call landed: customer satisfaction crossed with whether the agent
    itself was at fault, independent of any automated red-alert detector. Only counts calls
    we have both signals for (audited + full detail fetched)."""
    groups = {
        "all_clear": {"label": "All Clear", "description": "Customer was satisfied, agent performed correctly", "call_ids": []},
        "blind_spot": {"label": "Blind Spot", "description": "Customer was satisfied, but agent made mistakes", "call_ids": []},
        "false_alarm": {"label": "False Alarm", "description": "Customer was unhappy, but agent performed correctly", "call_ids": []},
        "red_alert": {"label": "Red Alert", "description": "Customer was unhappy, and agent made mistakes", "call_ids": []},
    }
    considered = 0
    for r in calls:
        if r.get("audit_verdict") is None or not has_call_detail(r):
            continue
        sat = call_satisfaction(r)
        if sat is None:
            continue
        considered += 1
        mistake = agent_made_mistake(r)
        key = ("all_clear" if (sat and not mistake) else
               "blind_spot" if (sat and mistake) else
               "false_alarm" if (not sat and not mistake) else
               "red_alert")
        groups[key]["call_ids"].append(r["id"])
    result = []
    for key in ("all_clear", "blind_spot", "false_alarm", "red_alert"):
        g = groups[key]
        count = len(g["call_ids"])
        result.append({
            "key": key,
            "label": g["label"],
            "description": g["description"],
            "count": count,
            "pct": round(count / considered * 100, 1) if considered else 0,
            "call_ids": g["call_ids"],
        })
    return {"considered": considered, "groups": result}


def compute_meta():
    all_calls = list(CALLS.values())
    if not all_calls:
        return {"total_calls": 0, "min_day": None, "max_day": None, "days_available": 0, "agents": []}
    days = sorted({day_key(r) for r in all_calls})
    agents = sorted({agent_key(r) for r in all_calls})
    return {
        "total_calls": len(all_calls),
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
        "from_number": rec.get("from_number"),
        "to_number": rec.get("to_number"),
        "end_reason": rec.get("end_reason"),
        "audit_verdict": rec.get("audit_verdict"),
        "cost_usd": rec.get("cost_usd"),
        "has_recording": rec.get("has_recording"),
    }


REQUEST_TIMEOUT_SECONDS = 10


def fetch_call_detail(call_id, api_key):
    if not api_key:
        return 500, json.dumps({"error": {"message": "Missing a SuperBryn API key for this call's agent", "code": "no_api_key"}}).encode()
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
        # Without this, a hung/unreachable upstream leaves the browser's "Loading call…" spinning forever.
        return 504, json.dumps({"error": {"message": f"Upstream request failed: {e.reason}", "code": "upstream_unreachable"}}).encode()


def fetch_audio_redirect(call_id, api_key):
    if not api_key:
        return 500, None
    conn = http.client.HTTPSConnection(API_HOST, timeout=REQUEST_TIMEOUT_SECONDS)
    try:
        conn.request(
            "GET",
            f"{BASE_PATH}/observability/calls/{call_id}/audio",
            headers={"X-API-Key": api_key},
        )
        resp = conn.getresponse()
        location = resp.getheader("Location")
        status = resp.status
        resp.read()
        return status, location
    except OSError:
        return 504, None
    finally:
        conn.close()


class Handler(BaseHTTPRequestHandler):
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

    def _serve_file(self, filename, content_type):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        with open(path, "rb") as f:
            self._send_bytes(f.read(), 200, content_type)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)
        days_param = qs.get("days", [None])[0]
        agent_param = qs.get("agent", [None])[0]

        if path in ("/", "/index.html"):
            self._serve_file("dashboard.html", "text/html")
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
        if path == "/api/agents":
            self._send_json(compute_agents(filtered_calls(days_param, agent_param)))
            return
        if path == "/api/lanes":
            self._send_json(compute_lanes(filtered_calls(days_param, agent_param)))
            return
        if path == "/api/issues":
            self._send_json(compute_issues(filtered_calls(days_param, agent_param)))
            return
        if path == "/api/quadrants":
            self._send_json(compute_quadrants(filtered_calls(days_param, agent_param)))
            return
        if path == "/api/issues/rca":
            self._send_json(list(ISSUE_RCA.values()))
            return
        if path == "/api/themes":
            self._send_json(THEMES)
            return
        if path == "/api/calls" and "ids" in qs:
            ids = qs["ids"][0].split(",")
            briefs = [call_brief(CALLS[i]) for i in ids if i in CALLS]
            self._send_json(briefs)
            return

        m = CALL_AUDIO_RE.match(path)
        if m:
            call_id = m.group(1)
            rec = CALLS.get(call_id)
            key = api_key_for(agent_key(rec) if rec else None)
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
            # Always live - no local transcript cache to keep in sync. Only call_id (plus the
            # list-view record used for issue/theme aggregation) is kept in CALLS; the full
            # detail (transcript/audio/metrics) for one specific call is fetched fresh on click.
            call_id = m.group(1)
            rec = CALLS.get(call_id)
            key = api_key_for(agent_key(rec) if rec else None)
            status, body = fetch_call_detail(call_id, key)
            self._send_bytes(body, status, "application/json")
            return

        self._send_json({"error": "not found"}, status=404)


def main():
    global DEFAULT_API_KEY

    parser = argparse.ArgumentParser(description="Run the SuperBryn call quality dashboard")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument(
        "--data",
        type=str,
        default=None,
        help="Path to a single fetched observability_calls_*.json file (default: merge every file under data/)",
    )
    args = parser.parse_args()

    DEFAULT_API_KEY = os.environ.get("SUPERBRYN_API_KEY")
    if not DEFAULT_API_KEY and not any(k.startswith("SUPERBRYN_API_KEY_") for k in os.environ):
        print(
            "Warning: no SUPERBRYN_API_KEY or SUPERBRYN_API_KEY_<AGENT> env var set - "
            "call detail/audio will fail to load live, but aggregate views still work.",
            file=sys.stderr,
        )

    base = os.path.dirname(os.path.abspath(__file__))

    if args.data:
        load_dataset(args.data)
    else:
        candidates = sorted(glob.glob(os.path.join(base, "data", "observability_calls_*.json")))
        if not candidates:
            sys.exit("No data file found under data/. Run fetch_superbryn_calls.py first.")
        for path in candidates:
            load_dataset(path)

    for path in sorted(glob.glob(os.path.join(base, "data", "red_alert_transcripts_*.json"))):
        load_call_detail(path)
    for path in sorted(glob.glob(os.path.join(base, "data", "audited_clean_transcripts*.json"))):
        load_call_detail(path)
    for path in sorted(glob.glob(os.path.join(base, "data", "legacy_week_transcripts*.json"))):
        load_call_detail(path)

    load_issue_rca(os.path.join(base, "data", "issue_rca.json"))
    load_root_cause_themes(os.path.join(base, "data", "root_cause_themes.json"))

    server = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    print(f"Dashboard running at http://127.0.0.1:{args.port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
