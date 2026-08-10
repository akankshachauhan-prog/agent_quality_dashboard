#!/usr/bin/env python3
"""Group red-alert calls by issue code and sample transcripts for RCA analysis.

analysis_issues (with per-call "evidence" text) is currently only populated on
data/observability_calls_sales_outbound_2026-07-21_to_2026-07-27.json - other
loaded datasets (sales_inbound, the newer sales_outbound fetch) have
analysis_issues: null on every red-alert call, so they're excluded here rather
than silently contributing empty buckets.

For each of the top issue codes in that file, this samples calls and pulls the
per-issue "evidence" note (already a targeted "why was this flagged" quote) plus
the full transcript (from red_alert_transcripts_sales_outbound_20260721.json,
fetched specifically for this window since the existing red_alert_transcripts_*
files don't cover these call ids).

The output (data/issue_bucket_samples.json) is meant to be read by a human/agent
to synthesize root-cause summaries into data/issue_rca.json - this script does no
LLM calls itself.

Usage:
    python3 build_issue_bucket_samples.py
    python3 build_issue_bucket_samples.py --top 15 --sample-size 18
"""
import argparse
import json
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")

CALLS_PATH = os.path.join(DATA_DIR, "observability_calls_sales_outbound_2026-07-21_to_2026-07-27.json")
TRANSCRIPTS_PATH = os.path.join(DATA_DIR, "red_alert_transcripts_sales_outbound_20260721.json")


def is_red_alert(rec):
    return rec.get("audit_verdict") in ("TP", "FN")


def load_calls(path):
    with open(path) as f:
        payload = json.load(f)
    return {rec["id"]: rec for rec in payload.get("data", [])}


def load_transcripts(path):
    if not os.path.exists(path):
        return {}
    with open(path) as f:
        payload = json.load(f)
    return {rec["id"]: rec for rec in payload.get("data", [])}


def bucket_issues(calls):
    """Groups (call_id, issue) pairs by code, keeping each issue's own evidence
    text - a call with N issues contributes to N buckets, each with its own
    evidence string for that specific flag."""
    groups = defaultdict(lambda: {"title": None, "lane": None, "code": None, "entries": []})
    for rec in calls.values():
        if not is_red_alert(rec):
            continue
        for issue in rec.get("analysis_issues") or []:
            g = groups[issue["code"]]
            g["title"] = issue["title"]
            g["lane"] = issue["lane"]
            g["code"] = issue["code"]
            g["entries"].append({"call_id": rec["id"], "evidence": issue.get("evidence")})
    return groups


def even_sample(entries, n):
    """Evenly spaced sample across the bucket rather than just the first n,
    so the sample isn't dominated by whichever fetch happened to list first."""
    if len(entries) <= n:
        return list(entries)
    step = len(entries) / n
    return [entries[int(i * step)] for i in range(n)]


def summarize_call(rec, transcript_rec, evidence):
    out = {
        "id": rec["id"],
        "duration_seconds": rec.get("duration_seconds"),
        "end_reason": rec.get("end_reason"),
        "audit_verdict": rec.get("audit_verdict"),
        "evidence": evidence,
    }
    if transcript_rec:
        out["transcript"] = [
            {"speaker": t.get("speaker"), "text": t.get("text")}
            for t in (transcript_rec.get("transcript") or [])
        ]
        audit = transcript_rec.get("audit") or {}
        user_intent = audit.get("user_intent") or {}
        out["audit_reason"] = user_intent.get("reason")
        out["observer_analysis"] = audit.get("observer_analysis")
        out["transcript_audit"] = audit.get("transcript_audit")
    return out


def main():
    parser = argparse.ArgumentParser(description="Sample red-alert calls per issue bucket for RCA")
    parser.add_argument("--top", type=int, default=15, help="Number of top issue buckets to sample")
    parser.add_argument("--sample-size", type=int, default=18, help="Calls to sample per bucket")
    parser.add_argument("--out", type=str, default=os.path.join(DATA_DIR, "issue_bucket_samples.json"))
    args = parser.parse_args()

    calls = load_calls(CALLS_PATH)
    transcripts = load_transcripts(TRANSCRIPTS_PATH)
    groups = bucket_issues(calls)

    buckets = sorted(groups.values(), key=lambda g: -len(g["entries"]))[: args.top]

    output = []
    for g in buckets:
        sampled = even_sample(g["entries"], args.sample_size)
        sample_calls = [
            summarize_call(calls[e["call_id"]], transcripts.get(e["call_id"]), e["evidence"]) for e in sampled
        ]
        missing_transcripts = sum(1 for c in sample_calls if "transcript" not in c)
        output.append(
            {
                "code": g["code"],
                "title": g["title"],
                "lane": g["lane"],
                "total_count": len(g["entries"]),
                "sample_size": len(sample_calls),
                "missing_transcripts": missing_transcripts,
                "sample_calls": sample_calls,
            }
        )

    with open(args.out, "w") as f:
        json.dump(output, f, indent=2)

    print(f"Loaded {len(calls)} calls, {len(transcripts)} transcripts.")
    for b in output:
        print(f"  {b['code']:<35} total={b['total_count']:<5} sampled={b['sample_size']:<3} missing_transcripts={b['missing_transcripts']}")
    print(f"Saved {len(output)} buckets to {args.out}")


if __name__ == "__main__":
    main()
