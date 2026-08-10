#!/usr/bin/env python3
"""Fetch full call detail (incl. transcript) for audited calls from SuperBryn.

Reads call ids out of the cached data/observability_calls_*.json files (produced
by fetch_superbryn_calls.py), fetches the full detail record for each matching
call from the SuperBryn API, and writes them to a single JSON file.

Defaults to red-alert calls only (audit_verdict TP/FN, per is_red_alert() in
server.py), since that's the highest-value subset for issue/root-cause drill-down.
Pass --verdicts to fetch a different slice - e.g. FP,TN to get the rest of the
audited population (needed for anything that requires the full audited set, like
a satisfaction x agent-correctness breakdown across every reviewed call, not just
the ones already confirmed bad).

Usage:
    export SUPERBRYN_API_KEY=sbryn_...
    python3 fetch_red_alert_transcripts.py
    python3 fetch_red_alert_transcripts.py --data-glob "data/observability_calls_sales_outbound_*.json"
    python3 fetch_red_alert_transcripts.py --out data/red_alert_transcripts.json
    python3 fetch_red_alert_transcripts.py --verdicts FP,TN --out data/audited_clean_transcripts.json
"""
import argparse
import glob
import json
import os
import sys
import time
import urllib.error
import urllib.request

API_HOST = "api.superbryn.com"
BASE_PATH = "/public-api/v1"


def load_calls(data_glob):
    calls = {}
    for path in sorted(glob.glob(data_glob)):
        with open(path) as f:
            payload = json.load(f)
        for rec in payload.get("data", []):
            calls[rec["id"]] = rec
    return calls


def fetch_call_detail(call_id, api_key, timeout=10):
    url = f"https://{API_HOST}{BASE_PATH}/observability/calls/{call_id}"
    req = urllib.request.Request(url, headers={"X-API-Key": api_key})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"call {call_id} detail fetch failed ({e.code}): {e.read().decode()}")
    except urllib.error.URLError as e:
        # Without a timeout, one unresponsive call hangs the whole run forever (this is what
        # happened - a run stalled for hours on a single call and never wrote any output).
        raise RuntimeError(f"call {call_id} detail fetch failed (network): {e.reason}")


def main():
    parser = argparse.ArgumentParser(description="Fetch full transcripts for red-alert calls")
    parser.add_argument(
        "--data-glob",
        type=str,
        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "observability_calls_*.json"),
        help="Glob of cached list-view JSON files to source red-alert call ids from",
    )
    parser.add_argument(
        "--out",
        type=str,
        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "red_alert_transcripts.json"),
    )
    parser.add_argument("--api-key", type=str, default=os.environ.get("SUPERBRYN_API_KEY"))
    parser.add_argument("--sleep", type=float, default=0.1, help="Seconds to sleep between requests")
    parser.add_argument(
        "--verdicts", type=str, default="TP,FN",
        help="Comma-separated audit_verdict values to fetch full detail for (default: TP,FN - red alerts)",
    )
    args = parser.parse_args()

    if not args.api_key:
        sys.exit("Missing API key. Set SUPERBRYN_API_KEY env var or pass --api-key.")

    verdicts = set(v.strip() for v in args.verdicts.split(","))
    calls = load_calls(args.data_glob)
    target_ids = [call_id for call_id, rec in calls.items() if rec.get("audit_verdict") in verdicts]
    print(f"Loaded {len(calls)} cached calls, {len(target_ids)} match verdicts {sorted(verdicts)}.")

    results = []
    errors = []
    for i, call_id in enumerate(target_ids, 1):
        try:
            detail = fetch_call_detail(call_id, args.api_key)
            results.append(detail)
            print(f"  [{i}/{len(target_ids)}] fetched {call_id} "
                  f"(transcript turns: {len(detail.get('transcript') or [])})")
        except RuntimeError as e:
            errors.append(str(e))
            print(f"  [{i}/{len(target_ids)}] error: {e}")
        time.sleep(args.sleep)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w") as f:
        json.dump({"count": len(results), "errors": errors, "data": results}, f, indent=2)

    print(f"Saved {len(results)} call transcripts to {args.out}")
    if errors:
        print(f"{len(errors)} calls failed to fetch (see 'errors' in output file).")


if __name__ == "__main__":
    main()
