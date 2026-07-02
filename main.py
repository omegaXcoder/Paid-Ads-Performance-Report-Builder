"""
Daily Ads Performance Report Builder
--------------------------------------------------------------------
Joins Ads_Raw_Metrics + WhatConverts_Raw_Leads (via Account_Mapping) into
your Google Sheets report template, writing one dated tab per day (plus a
"Previous Year" tab) inside a monthly report file — e.g.
"Ads Performance Reports - July 2026" containing tabs "July 2, 2026" and
"July 2, 2026 Previous Year".

Each day's numbers, including the Target-CPL comparison used for color
coding, are FROZEN as plain values at write time — never live formulas —
so later changes to a client's Target CPL never retroactively change how
a past day is displayed.

Defaults to a dry run: logs exactly what it would do, creates/writes
nothing, unless run with --live. Always emails a log either way.

See README.md for full usage.
"""

import argparse
import datetime
import json
import os
import re
import smtplib
import sys
from email.mime.text import MIMEText

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────

RAW_SPREADSHEET_ID = os.environ.get(
    "RAW_SPREADSHEET_ID", "1v9pqP0IQPsHLF45pTqlHvzkTtg1fHTYNVuNvSzbOU54"
).strip()
TEMPLATE_SPREADSHEET_ID = os.environ.get("TEMPLATE_SPREADSHEET_ID", "").strip()

# Leave blank to auto-detect: uses the same Drive folder that
# RAW_SPREADSHEET_ID already lives in.
REPORTS_FOLDER_ID = os.environ.get("REPORTS_FOLDER_ID", "").strip()

ADS_TAB = "Ads_Raw_Metrics"
WC_TAB = "WhatConverts_Raw_Leads"
MAPPING_TAB = "Account_Mapping"

TEMPLATE_CURRENT_TAB = "Dashboard"
TEMPLATE_PRIOR_YEAR_TAB = "Last Year Performance"
TEMPLATE_MAPPING_TAB = "Account Mapping (Source)"

EMAIL_TO = os.environ.get("EMAIL_TO", "omega@kudos.marketing")
EMAIL_FROM = os.environ.get("EMAIL_FROM", "automation@kudos.marketing")

# Row ranges are fixed by the template's layout — data rows then a TOTAL
# row, for each of the 3 period tables.
PERIOD_ROWS = {
    "Last 7 Days": {"data": (7, 16), "total": 17, "section_cell": "A5", "snapshot_cell": "A3"},
    "Previous Full Week (Mon-Sun)": {"data": (21, 30), "total": 31, "section_cell": "A19"},
    "Month to Date": {"data": (35, 44), "total": 45, "section_cell": "A33"},
}
PERIOD_ORDER = ["Last 7 Days", "Previous Full Week (Mon-Sun)", "Month to Date"]

FOOTNOTE_CELL = "A47"


# ── Google API clients ───────────────────────────────────────────────────

def get_clients():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    raw_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw_json:
        info = json.loads(raw_json)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        file_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
        creds = Credentials.from_service_account_file(file_path, scopes=scopes)

    gc = gspread.authorize(creds)
    drive_service = build("drive", "v3", credentials=creds)
    # service_account_email is not sensitive (it's a public identifier, not
    # a credential) — safe to log for debugging "who is actually running this."
    service_account_email = getattr(creds, "service_account_email", "unknown")
    return gc, drive_service, service_account_email


# ── Period label parsing ─────────────────────────────────────────────────

def parse_period(period_string):
    """
    Splits a raw Period string (e.g. 'Last 7 Days (2026-06-25 to 2026-07-01)'
    or 'Month to Date - Prior Year (2025-06-30 to 2025-06-30)') into
    (base_label, is_prior_year, date_range_str).
    """
    is_prior_year = " - Prior Year" in period_string

    base_label = None
    for label in PERIOD_ORDER:
        if period_string.startswith(label):
            base_label = label
            break

    match = re.search(r"\(([\d-]+ to [\d-]+)\)\s*$", period_string)
    date_range_str = match.group(1) if match else ""

    return base_label, is_prior_year, date_range_str


# ── Reading raw data ─────────────────────────────────────────────────────

def find_col(headers, *keywords):
    """Case-insensitive best-effort header match by keyword(s)."""
    for h in headers:
        low = h.lower()
        if all(k.lower() in low for k in keywords):
            return h
    return None


def read_account_mapping(raw_ss):
    ws = raw_ss.worksheet(MAPPING_TAB)
    rows = ws.get_all_records()
    headers = list(rows[0].keys()) if rows else []

    name_col = find_col(headers, "business", "name") or "Business Name"
    profile_col = find_col(headers, "profile", "id") or "What Converts Profile ID"
    customer_col = find_col(headers, "customer", "id") or "Google Ads Customer ID"
    target_col = find_col(headers, "target") or find_col(headers, "cost", "qualified")

    mapping = []
    for row in rows:
        name = str(row.get(name_col, "")).strip()
        profile_id = str(row.get(profile_col, "")).strip()
        customer_id = str(row.get(customer_col, "")).strip()
        target_raw = row.get(target_col, "") if target_col else ""
        try:
            target_cpl = float(target_raw) if str(target_raw).strip() != "" else None
        except (TypeError, ValueError):
            target_cpl = None

        if name:
            mapping.append({
                "name": name,
                "profile_id": profile_id,
                "customer_id": customer_id,
                "target_cpl": target_cpl,
            })
    return mapping


def read_ads_raw(raw_ss):
    ws = raw_ss.worksheet(ADS_TAB)
    rows = ws.get_all_records()

    parsed = []
    for row in rows:
        base_label, is_prior_year, date_range = parse_period(str(row.get("Period", "")))
        if not base_label:
            continue
        parsed.append({
            "customer_id": str(row.get("Account ID", "")).strip(),
            "base_label": base_label,
            "is_prior_year": is_prior_year,
            "date_range": date_range,
            "impressions": float(row.get("Impressions") or 0),
            "clicks": float(row.get("Clicks") or 0),
            "cost": float(row.get("Cost") or 0),
            "ctr": float(row.get("CTR") or 0),
            "conversions": float(row.get("Conversions") or 0),
            "cost_per_conversion": row.get("Cost Per Conversion"),
        })
    return parsed


def read_wc_raw(raw_ss):
    ws = raw_ss.worksheet(WC_TAB)
    rows = ws.get_all_records()

    parsed = []
    for row in rows:
        base_label, is_prior_year, date_range = parse_period(str(row.get("Period", "")))
        if not base_label:
            continue
        parsed.append({
            "profile_id": str(row.get("WhatConverts Profile ID", "")).strip(),
            "base_label": base_label,
            "is_prior_year": is_prior_year,
            "date_range": date_range,
            "qualified": float(row.get("Qualified Leads") or 0),
            "quote_value": float(row.get("Qualified Quote Value") or 0),
            "sales_value": float(row.get("Qualified Sales Value") or 0),
        })
    return parsed


# ── Joining ────────────────────────────────────────────────────────────

def build_joined_data(mapping, ads_rows, wc_rows):
    """
    Returns data[is_prior_year][base_label][business_name] = {...merged...}
    Joined via Account_Mapping: Ads rows match on customer_id,
    WhatConverts rows match on profile_id.
    """
    ads_index = {(r["customer_id"], r["base_label"], r["is_prior_year"]): r for r in ads_rows}
    wc_index = {(r["profile_id"], r["base_label"], r["is_prior_year"]): r for r in wc_rows}

    data = {False: {p: {} for p in PERIOD_ORDER}, True: {p: {} for p in PERIOD_ORDER}}

    for client in mapping:
        for is_py in (False, True):
            for period in PERIOD_ORDER:
                ads = ads_index.get((client["customer_id"], period, is_py))
                wc = wc_index.get((client["profile_id"], period, is_py))

                spend = ads["cost"] if ads else 0.0
                clicks = ads["clicks"] if ads else 0.0
                impressions = ads["impressions"] if ads else 0.0
                ctr = ads["ctr"] if ads else 0.0
                conversions = ads["conversions"] if ads else 0.0
                cost_per_conversion = ads["cost_per_conversion"] if ads else None

                qualified = wc["qualified"] if wc else 0.0
                sales_value = wc["sales_value"] if wc else 0.0

                cost_per_qual = (spend / qualified) if qualified > 0 else None
                roas = (sales_value / spend) if spend > 0 else None

                target_cpl = client["target_cpl"]
                pct_diff = None
                if target_cpl is not None and cost_per_qual is not None:
                    pct_diff = (cost_per_qual - target_cpl) / target_cpl

                data[is_py][period][client["name"]] = {
                    "spend": spend,
                    "clicks": clicks,
                    "impressions": impressions,
                    "ctr": ctr,
                    "conversions": conversions,
                    "cost_per_conversion": cost_per_conversion,
                    "qualified": qualified,
                    "sales_value": sales_value,
                    "cost_per_qual": cost_per_qual,
                    "roas": roas,
                    "target_cpl": target_cpl,
                    "pct_diff": pct_diff,
                    "spend_vs_prior_wk": None,
                    "qual_vs_prior_wk": None,
                    "date_range": (ads["date_range"] if ads else (wc["date_range"] if wc else "")),
                }

    # Trend columns: only for "Last 7 Days", compared against
    # "Previous Full Week (Mon-Sun)" within the same is_prior_year bucket.
    for is_py in (False, True):
        last7 = data[is_py]["Last 7 Days"]
        prevwk = data[is_py]["Previous Full Week (Mon-Sun)"]
        for name, row in last7.items():
            prev = prevwk.get(name)
            if not prev:
                continue
            if prev["spend"] > 0:
                row["spend_vs_prior_wk"] = (row["spend"] - prev["spend"]) / prev["spend"]
            if prev["qualified"] > 0:
                row["qual_vs_prior_wk"] = (row["qualified"] - prev["qualified"]) / prev["qualified"]

    return data


def compute_total_row(period_data):
    """period_data: dict of business_name -> row dict for one period/year bucket."""
    rows = list(period_data.values())
    spend = sum(r["spend"] for r in rows)
    clicks = sum(r["clicks"] for r in rows)
    impressions = sum(r["impressions"] for r in rows)
    conversions = sum(r["conversions"] for r in rows)
    qualified = sum(r["qualified"] for r in rows)
    sales_value = sum(r["sales_value"] for r in rows)

    ctr = (clicks / impressions) if impressions > 0 else 0
    cost_per_qual = (spend / qualified) if qualified > 0 else None
    cost_per_conversion = (spend / conversions) if conversions > 0 else None
    roas = (sales_value / spend) if spend > 0 else 0

    return {
        "spend": round(spend, 2),
        "clicks": clicks,
        "ctr": round(ctr, 4),
        "qualified": qualified,
        "cost_per_qual": round(cost_per_qual, 2) if cost_per_qual is not None else "",
        "cost_per_conversion": round(cost_per_conversion, 2) if cost_per_conversion is not None else "",
        "sales_value": round(sales_value, 2),
        "roas": round(roas, 4),
    }


# ── Drive / Sheets file management ──────────────────────────────────────

def get_reports_folder_id(drive_service):
    if REPORTS_FOLDER_ID:
        return REPORTS_FOLDER_ID
    file = drive_service.files().get(fileId=RAW_SPREADSHEET_ID, fields="parents").execute()
    parents = file.get("parents", [])
    if not parents:
        raise RuntimeError("Could not determine parent folder of RAW_SPREADSHEET_ID.")
    return parents[0]


def find_or_create_month_file(gc, drive_service, folder_id, file_name, log, dry_run):
    query = (
        f"name = '{file_name}' and '{folder_id}' in parents "
        f"and trashed = false and mimeType = 'application/vnd.google-apps.spreadsheet'"
    )
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])

    if files:
        log(f"Found existing month file: {file_name} ({files[0]['id']})")
        return gc.open_by_key(files[0]["id"]), False

    log(f"Month file not found.")
    if dry_run:
        log(f"DRY RUN — would create: {file_name}")
        return None, True

    log(f"Creating new month file: {file_name}")

    if not TEMPLATE_SPREADSHEET_ID:
        raise RuntimeError("TEMPLATE_SPREADSHEET_ID is not set — cannot create a new month file.")

    copied = drive_service.files().copy(
        fileId=TEMPLATE_SPREADSHEET_ID,
        body={"name": file_name, "parents": [folder_id]},
    ).execute()
    log(f"Created new month file: {file_name} ({copied['id']})")
    return gc.open_by_key(copied["id"]), True


def seed_account_mapping_tab(month_ss, mapping, log, dry_run):
    log(f"Seeding {TEMPLATE_MAPPING_TAB} with {len(mapping)} clients (one-time, new month only).")
    if dry_run:
        return
    ws = month_ss.worksheet(TEMPLATE_MAPPING_TAB)
    values = [[m["name"], m["profile_id"], m["customer_id"], m["target_cpl"]] for m in mapping]
    ws.update(f"A5:D{4 + len(values)}", values, value_input_option="USER_ENTERED")


# ── Writing a day's data into a tab ──────────────────────────────────────

def duplicate_and_rename(month_ss, source_tab_name, new_tab_name, log, dry_run):
    log(f"Duplicating '{source_tab_name}' -> '{new_tab_name}'")
    if dry_run:
        return None
    source_ws = month_ss.worksheet(source_tab_name)
    new_ws = month_ss.duplicate_sheet(source_ws.id, new_sheet_name=new_tab_name)
    return new_ws


def write_period_table(ws, period, business_order, period_data, snapshot_generated_at, log):
    rows_cfg = PERIOD_ROWS[period]
    start_row, end_row = rows_cfg["data"]
    total_row = rows_cfg["total"]

    # Section header text
    date_range = ""
    for row in period_data.values():
        if row["date_range"]:
            date_range = row["date_range"]
            break
    section_text = f"Account Performance — {period} ({date_range})"
    ws.update_acell(rows_cfg["section_cell"], section_text)

    if "snapshot_cell" in rows_cfg:
        snapshot_text = (
            f"Snapshot: {period} ({date_range})  |  Data sources: Google Ads + WhatConverts  |  "
            f"Generated {snapshot_generated_at}"
        )
        ws.update_acell(rows_cfg["snapshot_cell"], snapshot_text)

    # Data rows
    data_matrix = []
    hidden_matrix = []  # M, N columns
    for name in business_order:
        r = period_data.get(name)
        if not r:
            data_matrix.append([name, "", "", "", "", "", "", "", "", "", ""])
            hidden_matrix.append(["", ""])
            continue

        data_matrix.append([
            name,
            round(r["spend"], 2),
            r["clicks"],
            round(r["ctr"], 4),
            r["qualified"],
            round(r["cost_per_qual"], 2) if r["cost_per_qual"] is not None else "",
            round(r["cost_per_conversion"], 2) if r["cost_per_conversion"] not in (None, "") else "",
            round(r["sales_value"], 2),
            round(r["roas"], 4) if r["roas"] is not None else "",
            round(r["spend_vs_prior_wk"], 4) if r["spend_vs_prior_wk"] is not None else "",
            round(r["qual_vs_prior_wk"], 4) if r["qual_vs_prior_wk"] is not None else "",
        ])
        hidden_matrix.append([
            r["target_cpl"] if r["target_cpl"] is not None else "",
            round(r["pct_diff"], 4) if r["pct_diff"] is not None else "",
        ])

    ws.update(f"A{start_row}:K{end_row}", data_matrix, value_input_option="USER_ENTERED")
    ws.update(f"M{start_row}:N{end_row}", hidden_matrix, value_input_option="USER_ENTERED")

    # TOTAL / BLENDED row
    totals = compute_total_row(period_data)
    ws.update(f"A{total_row}:I{total_row}", [[
        "TOTAL / BLENDED", totals["spend"], totals["clicks"], totals["ctr"],
        totals["qualified"], totals["cost_per_qual"], totals["cost_per_conversion"],
        totals["sales_value"], totals["roas"],
    ]], value_input_option="USER_ENTERED")

    log(f"Wrote {period} table ({len(business_order)} clients + total row).")


def clear_footnote(ws, log):
    ws.update_acell(FOOTNOTE_CELL, "")
    log(f"Cleared footnote cell {FOOTNOTE_CELL}.")


# ── Email log ────────────────────────────────────────────────────────────

def send_log_email(subject, body):
    host = os.environ.get("SMTP_HOST")
    port = int(os.environ.get("SMTP_PORT", "587"))
    username = os.environ.get("SMTP_USERNAME")
    password = os.environ.get("SMTP_PASSWORD")

    if not (host and username and password and EMAIL_TO):
        print("SMTP not fully configured — skipping email, log was still printed above.")
        return

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO

    with smtplib.SMTP(host, port) as server:
        server.starttls()
        server.login(username, password)
        server.sendmail(EMAIL_FROM, [EMAIL_TO], msg.as_string())


# ── Main ─────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Daily Ads performance report builder")
    parser.add_argument("--live", action="store_true", help="Actually create/write. Default is dry run.")
    args = parser.parse_args()
    dry_run = not args.live or os.environ.get("DRY_RUN", "").lower() == "true"

    log_lines = []

    def log(msg):
        line = f"{datetime.datetime.now().isoformat()} {msg}"
        print(line)
        log_lines.append(line)

    log(f"=== Starting run | mode={'DRY RUN' if dry_run else 'LIVE'} ===")

    today = datetime.date.today()
    month_file_name = f"Ads Performance Reports - {today.strftime('%B %Y')}"
    day_tab_name = today.strftime("%B %-d, %Y")
    prior_year_tab_name = f"{day_tab_name} Previous Year"
    generated_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    errors = []
    try:
        gc, drive_service, service_account_email = get_clients()
        log(f"Authenticated as service account: {service_account_email}")
        log(f"TEMPLATE_SPREADSHEET_ID length: {len(TEMPLATE_SPREADSHEET_ID)} "
            f"(should be ~44 chars, no spaces or slashes)")
        if not TEMPLATE_SPREADSHEET_ID:
            raise RuntimeError(
                "TEMPLATE_SPREADSHEET_ID is empty — check the GitHub secret is set and saved."
            )

        raw_ss = gc.open_by_key(RAW_SPREADSHEET_ID)
        mapping = read_account_mapping(raw_ss)
        ads_rows = read_ads_raw(raw_ss)
        wc_rows = read_wc_raw(raw_ss)
        log(f"Loaded {len(mapping)} clients, {len(ads_rows)} Ads rows, {len(wc_rows)} WhatConverts rows.")

        joined = build_joined_data(mapping, ads_rows, wc_rows)
        business_order = [m["name"] for m in mapping]

        folder_id = get_reports_folder_id(drive_service)
        month_ss, is_new = find_or_create_month_file(
            gc, drive_service, folder_id, month_file_name, log, dry_run
        )

        if is_new and not dry_run:
            seed_account_mapping_tab(month_ss, mapping, log, dry_run)
        elif is_new and dry_run:
            seed_account_mapping_tab(None, mapping, log, dry_run)

        for is_py, tab_source, tab_name in (
            (False, TEMPLATE_CURRENT_TAB, day_tab_name),
            (True, TEMPLATE_PRIOR_YEAR_TAB, prior_year_tab_name),
        ):
            ws = duplicate_and_rename(month_ss, tab_source, tab_name, log, dry_run)
            if dry_run:
                for period in PERIOD_ORDER:
                    log(f"  Would write {period} ({'Prior Year' if is_py else 'Current'}) "
                        f"for {len(business_order)} clients.")
                continue

            for period in PERIOD_ORDER:
                write_period_table(ws, period, business_order, joined[is_py][period], generated_at, log)
            clear_footnote(ws, log)

        log(f"Processed {len(business_order)} clients across {len(PERIOD_ORDER)} periods x 2 (current/prior year).")

    except Exception as e:
        errors.append(str(e))
        log(f"FATAL ERROR: {e}")

    mode_label = "DRY RUN" if dry_run else ("COMPLETED WITH ERRORS" if errors else "SUCCESS")
    subject = f"AUTOMATION LOGGING: Ads Performance Report Builder — {mode_label}"
    body = "\n".join(log_lines)
    if errors:
        body += "\n\nErrors:\n" + "\n".join(errors)
    send_log_email(subject, body)

    log("=== Run finished ===")

    if errors and not dry_run:
        sys.exit(1)


if __name__ == "__main__":
    main()
