#!/usr/bin/env python3
"""Build data/snapshot.json - a compact, privacy-safe copy of the locally fetched call data
that's actually committed to git and deployed to Vercel.

Strips everything the deployed dashboard doesn't need: transcript text, phone numbers,
customer names, and the ~40-70 pass/not_applicable metric_evaluations entries per call -
keeping only the flagged (actual) issues plus the handful of scalar fields the aggregate
endpoints (summary/days/issues) read. This is what makes it safe to commit: no conversation
content or PII ever leaves your machine, only issue codes/severities, verdicts, costs, and
timestamps.

This is also the "refresh" step: after re-running fetch_superbryn_calls.py and
fetch_red_alert_transcripts.py to pull newer calls, re-run this script, then commit +
push data/snapshot.json to redeploy with fresh aggregates. Per-call transcript/audio in the
live dashboard is fetched on demand from the SuperBryn API when a call is clicked - it isn't
part of this snapshot at all.

Usage:
    python3 build_snapshot.py
"""
import glob
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
OUT_PATH = os.path.join(DATA_DIR, "snapshot.json")


def load_merged_calls():
    calls = {}
    for path in sorted(glob.glob(os.path.join(DATA_DIR, "observability_calls_*.json"))):
        with open(path) as f:
            payload = json.load(f)
        for rec in payload["data"]:
            calls[rec["id"]] = rec
    detail_globs = ["red_alert_transcripts_*.json", "audited_clean_transcripts*.json", "legacy_week_transcripts*.json"]
    for pattern in detail_globs:
        for path in sorted(glob.glob(os.path.join(DATA_DIR, pattern))):
            with open(path) as f:
                payload = json.load(f)
            for rec in payload["data"]:
                calls[rec["id"]] = {**calls.get(rec["id"], {}), **rec}
    return calls


def call_issues(rec):
    if rec.get("analysis_issues"):
        return rec["analysis_issues"]
    return [e for e in (rec.get("metric_evaluations") or []) if e.get("result") == "flagged"]


def compact(rec, issue_defs):
    codes = []
    for i in call_issues(rec):
        issue_defs.setdefault(i["code"], {"title": i["title"], "lane": i["lane"], "severity": i.get("severity")})
        codes.append(i["code"])
    return {
        "id": rec["id"],
        "started_at": rec["started_at"],
        "duration_seconds": rec.get("duration_seconds"),
        "agent_label": rec.get("agent_label") or rec.get("agent_name") or "unknown",
        "cost_usd": rec.get("cost_usd"),
        "audit_verdict": rec.get("audit_verdict"),
        "has_recording": rec.get("has_recording"),
        "issue_codes": codes,
    }


def main():
    calls = load_merged_calls()
    issue_defs = {}
    snapshot = [compact(rec, issue_defs) for rec in calls.values()]
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(OUT_PATH, "w") as f:
        json.dump({"count": len(snapshot), "issue_defs": issue_defs, "data": snapshot}, f)
    size_kb = os.path.getsize(OUT_PATH) / 1024
    print(f"Wrote {OUT_PATH}: {len(snapshot)} calls, {len(issue_defs)} issue codes, {size_kb:.0f} KB")


if __name__ == "__main__":
    main()
