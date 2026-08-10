#!/usr/bin/env python3
"""Fetch full call detail for a specific list of call ids (not derived from a
cached list-view glob) - used to backfill detail for calls whose day range
wasn't covered by an earlier fetch_red_alert_transcripts.py run.

Usage:
    python3 fetch_specific_call_details.py --ids-file /tmp/ids.json --out data/out.json
"""
import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.request

API_HOST = "api.superbryn.com"
BASE_PATH = "/public-api/v1"


def fetch_call_detail(call_id, api_key):
    url = f"https://{API_HOST}{BASE_PATH}/observability/calls/{call_id}"
    req = urllib.request.Request(url, headers={"X-API-Key": api_key})
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"call {call_id} detail fetch failed ({e.code}): {e.read().decode()}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--ids-file", required=True, help="JSON file containing a list of call ids")
    parser.add_argument("--out", required=True)
    parser.add_argument("--api-key", default=os.environ.get("SUPERBRYN_API_KEY"))
    parser.add_argument("--sleep", type=float, default=0.1)
    args = parser.parse_args()

    if not args.api_key:
        sys.exit("Missing API key. Set SUPERBRYN_API_KEY env var or pass --api-key.")

    with open(args.ids_file) as f:
        ids = json.load(f)

    results, errors = [], []
    for i, call_id in enumerate(ids, 1):
        try:
            detail = fetch_call_detail(call_id, args.api_key)
            results.append(detail)
            print(f"  [{i}/{len(ids)}] fetched {call_id}")
        except RuntimeError as e:
            errors.append(str(e))
            print(f"  [{i}/{len(ids)}] error: {e}")
        time.sleep(args.sleep)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w") as f:
        json.dump({"count": len(results), "errors": errors, "data": results}, f, indent=2)
    print(f"Saved {len(results)} call details to {args.out}")


if __name__ == "__main__":
    main()
