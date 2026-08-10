#!/usr/bin/env python3
"""Build reports/quality_tracker_Aug2026.xlsx - a local mirror of the shared
Google Sheet quality tracker, with Ticket/Owner/Status/ETA columns per issue
row and one date column per audited day.

Day values are hardcoded here (pulled from the observability/red-alert data
for 1-5 Aug via server.py's compute_issues()-equivalent logic, and for 6-7 Aug
via build root_cause pipeline scripts in this repo) rather than recomputed,
since this mirrors a manually-curated tracker - update DAYS/SALES_INBOUND/
SALES_OUTBOUND below when adding a new day.

Usage:
    python3 build_quality_tracker_sheet.py
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(HERE, "reports", "quality_tracker_Aug2026.xlsx")

DAYS = ["1 Aug", "2 Aug", "3 Aug", "4 Aug", "5 Aug", "6 Aug", "7 Aug"]

JIRA_ROWS = [
    ("https://spyne.atlassian.net/browse/RETCONVAI-4406", "", "Piyush",
     "Already rolled out for IB agents/ in pipeline for OB Waiting for livekit to increase concurrency"),
    ("https://spyne.atlassian.net/browse/RETCONVAI-4407", "", "Bhaskar", "In progress"),
    ("https://spyne.atlassian.net/browse/RETCONVAI-4408", "", "Bhaskar",
     "already fixed and deployed on production for new service prompt"),
]

SALES_INBOUND = [
    ("Total calls", [132, 34, 290, 284, 266, 161, 0]),
    ("Audited calls", [70, 20, 120, 58, 123, 62, 0]),
    ("% Red Alerts", ["37%", "44%", "29%", "14%", "28%", "26%", "0%"]),
    ("Issue 1: Agent speaking too fast", [49, 15, 83, 41, 73, 42, 0]),
    ("Issue 2: Word error rate too high", [37, 14, 71, 38, 63, 38, 0]),
    ("Issue 3: Agent response time too slow", [37, 14, 67, 32, 58, 33, 0]),
    ("Issue 4: Agent not responding", [16, 2, 30, 11, 25, 22, 0]),
    ("Issue 5: Agent provides incorrect information", [2, 1, 3, 0, 1, 1, 0]),
    ("Issue 6: Agent misrepresented tool results", [1, 0, 4, 0, 3, 1, 0]),
    ("Issue 7: Agent re-asks already answered questions", [0, 1, 3, 1, 3, 1, 0]),
    ("Issue 8: Required lead information not captured or confirmed", [0, 2, 1, 1, 0, 0, 0]),
]

SALES_OUTBOUND = [
    ("Total calls", [1375, 9, 1915, 1163, 1879, 595, 0]),
    ("Audited calls", [90, 1, 148, 33, 190, 49, 0]),
    ("% Red Alerts", ["6%", "11%", "6%", "2%", "7%", "6%", "0%"]),
    ("Issue 1: Word error rate too high", [63, 1, 90, 24, 105, 22, 0]),
    ("Issue 2: Agent speaking too fast", [57, 1, 81, 24, 104, 24, 0]),
    ("Issue 3: Agent did not follow the expected call flow", [34, 0, 49, 16, 73, 14, 0]),
    ("Issue 4: Agent did not deliver the opening greeting", [19, 0, 32, 7, 41, 10, 0]),
    ("Issue 5: Agent response time too slow", [18, 0, 26, 7, 41, 9, 0]),
    ("Issue 6: Agent speaking too slow", [11, 0, 19, 5, 21, 3, 0]),
    ("Issue 7: IVR / recorded message answered", [16, 0, 10, 4, 36, 2, 0]),
    ("Issue 8: Call went to voicemail", [6, 0, 8, 0, 4, 0, 0]),
]

SERVICE_ROWS = ["Total calls", "Audited calls", "% Red Alerts", "Issue 1", "Issue 2", "Issue 3", "Issue 4", "Issue 5"]

HEADER_FILL = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NEW_DAY_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")  # 6-7 Aug: newly added
CORRECTED_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")  # 5 Aug: refetched/corrected
BOLD = Font(bold=True)


def write_section(ws, name, data_rows):
    r = ws.max_row + 1 if ws.max_row > 1 else 1
    header = [name, "Ticket", "Owner", "Status", "ETA"] + DAYS
    ws.append(header)
    for c in range(1, len(header) + 1):
        cell = ws.cell(row=ws.max_row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for label, vals in data_rows:
        row = [label, "", "", "", ""] + list(vals)
        ws.append(row)
        ws.cell(row=ws.max_row, column=1).font = BOLD
        # 5 Aug (corrected via refetch) is 3rd-from-last column; 6/7 Aug (newly added) are the last 2
        ws.cell(row=ws.max_row, column=len(header) - 2).fill = CORRECTED_FILL
        for c in (len(header) - 1, len(header)):
            ws.cell(row=ws.max_row, column=c).fill = NEW_DAY_FILL
    ws.append(["..."])
    ws.append([])


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Quality Tracker"

    ws.append(["JIRA EPIC for Quality issues", "Created date", "Owner", "Status"])
    for c in range(1, 5):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for ticket_url, created, owner, status in JIRA_ROWS:
        ws.append([ticket_url, created, owner, status])
    ws.append(["..."])
    ws.append([])

    write_section(ws, "Sales inbound", SALES_INBOUND)
    write_section(ws, "Sales outbound", SALES_OUTBOUND)
    write_section(ws, "Service inbound", [(label, [""] * 7) for label in SERVICE_ROWS])
    write_section(ws, "Service outbound", [(label, [""] * 7) for label in SERVICE_ROWS])

    ws.append([
        "Sales inbound/outbound figures pulled from observability + red-alert audit data (Aug 1-7, 2026). "
        "5 Aug (highlighted amber) was re-fetched and corrected - the original figures were captured while "
        "that day was still in progress (only ~16/105 calls in), so totals/audits/issue counts were far too "
        "low; the numbers here reflect the completed day. 6 Aug is a full day; 7 Aug is 0 across the board "
        "because no calls had landed yet at fetch time. Service inbound/outbound has no data pipeline yet - "
        "fill in manually."
    ])

    widths = [58, 14, 12, 40, 10] + [8] * len(DAYS)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
