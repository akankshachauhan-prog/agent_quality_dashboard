#!/usr/bin/env python3
"""Fetch SuperBryn observability/eval call data for a date range and save as JSON.

Usage:
    export SUPERBRYN_API_KEY=sbryn_...
    python3 fetch_superbryn_calls.py                # last 7 days, observability + evals
    python3 fetch_superbryn_calls.py --days 14
    python3 fetch_superbryn_calls.py --start 2026-07-01 --end 2026-07-27
    python3 fetch_superbryn_calls.py --label sales_outbound   # tag records + filename (avoids overwriting another agent's fetch)
"""
import argparse
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta

BASE_URL = "https://api.superbryn.com/public-api/v1"


def fetch_page(endpoint_path, api_key, params):
    url = f"{BASE_URL}/{endpoint_path}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={"X-API-Key": api_key})
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        raise RuntimeError(f"{endpoint_path} request failed ({e.code}): {body}")


def fetch_day(endpoint_path, api_key, day_str, tz, limit=200):
    records = []
    cursor = None
    while True:
        params = {"date": day_str, "limit": limit}
        if cursor:
            params["cursor"] = cursor
        page = fetch_page(endpoint_path, api_key, params)
        records.extend(page.get("data", []))
        if page.get("has_more") and page.get("next_cursor"):
            cursor = page["next_cursor"]
        else:
            break
    return records


def daterange(start, end):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def main():
    parser = argparse.ArgumentParser(description="Fetch SuperBryn call data for a date range")
    parser.add_argument("--days", type=int, default=7, help="Trailing days to fetch, ending on --end (default 7)")
    parser.add_argument("--start", type=str, help="Start date YYYY-MM-DD (overrides --days)")
    parser.add_argument("--end", type=str, help="End date YYYY-MM-DD (default: today)")
    parser.add_argument("--tz", type=str, default="Asia/Kolkata", help="IANA timezone for day boundaries")
    parser.add_argument(
        "--endpoints", nargs="+", default=["observability", "evals"], choices=["observability", "evals"]
    )
    parser.add_argument("--out-dir", type=str, default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "data"))
    parser.add_argument("--api-key", type=str, default=os.environ.get("SUPERBRYN_API_KEY"))
    parser.add_argument(
        "--label",
        type=str,
        default=None,
        help="Tag for this fetch (e.g. an agent name). Stamped onto each record as agent_label and appended "
        "to the output filename, so fetching a second API key doesn't overwrite the first one's file.",
    )
    args = parser.parse_args()

    if not args.api_key:
        sys.exit("Missing API key. Set SUPERBRYN_API_KEY env var or pass --api-key.")

    end = date.fromisoformat(args.end) if args.end else date.today()
    start = date.fromisoformat(args.start) if args.start else end - timedelta(days=args.days - 1)

    os.makedirs(args.out_dir, exist_ok=True)

    for endpoint_name in args.endpoints:
        endpoint_path = f"{endpoint_name}/calls"
        all_records = []
        for day in daterange(start, end):
            day_str = day.isoformat()
            print(f"Fetching {endpoint_name} calls for {day_str} ({args.tz})...")
            try:
                records = fetch_day(endpoint_path, args.api_key, day_str, args.tz)
            except RuntimeError as e:
                print(f"  error: {e}")
                continue
            print(f"  {len(records)} records")
            if args.label:
                for r in records:
                    r["agent_label"] = args.label
            all_records.extend(records)

        label_suffix = f"{args.label}_" if args.label else ""
        out_file = os.path.join(
            args.out_dir, f"{endpoint_name}_calls_{label_suffix}{start.isoformat()}_to_{end.isoformat()}.json"
        )
        with open(out_file, "w") as f:
            json.dump(
                {
                    "endpoint": endpoint_path,
                    "start_date": start.isoformat(),
                    "end_date": end.isoformat(),
                    "tz": args.tz,
                    "fetched_at": datetime.utcnow().isoformat() + "Z",
                    "count": len(all_records),
                    "data": all_records,
                },
                f,
                indent=2,
            )
        print(f"Saved {len(all_records)} {endpoint_name} records to {out_file}")


if __name__ == "__main__":
    main()
