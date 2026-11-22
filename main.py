import hashlib
import json
import logging
import math
import os
import re
import secrets
import sys
import time
import uuid
from datetime import date, datetime, timedelta
from urllib.parse import parse_qs, urlsplit
from zoneinfo import ZoneInfo

import gspread
import requests
from google.oauth2.service_account import Credentials


# ============================================================
# 基本設定
# ============================================================

META_API_VERSION = "v27.0"

JST = ZoneInfo("Asia/Tokyo")

DEFAULT_WORKSHEET_NAME = "gitreport"

# Webサイト上のLeadイベント
TARGET_ACTION_TYPE = "offsite_conversion.fb_pixel_lead"

# Google Sheetsの日付シリアル値の基準日
GOOGLE_SHEETS_EPOCH = date(1899, 12, 30)


# ============================================================
# main
# ============================================================

class ExportError(Exception):
    """Error whose message can be printed without exposing credentials."""


def main():
    print("=== Start Meta Lead Export ===")

    config = load_secret()
    resolved = resolve_config(config)
    validate_config(resolved)

    dest = resolved["sheet"]
    book = connect_spreadsheet(dest["spreadsheet_id"], dest["google_service_account"])
    lease = acquire_lease(book, dest["worksheet_name"])
    if lease is None:
        print("Export skipped: another run owns the destination.")
        return

    try:
        monthly_12_ranges, monthly_2_ranges, daily_since, daily_until = get_target_date_ranges()
        rows = fetch_meta_rows(
            act_id=resolved["meta"]["account_id"],
            token=resolved["meta"]["token"],
            monthly_12_ranges=monthly_12_ranges,
            monthly_2_ranges=monthly_2_ranges,
            daily_since=daily_since,
            daily_until=daily_until,
        )
        write_to_sheet(book, dest["worksheet_name"], sort_rows(rows), lease)
    except BaseException:
        release_lease(book, lease)
        raise

    print(f"Total rows written: {len(rows)}")
    print("=== Completed ===")


# ============================================================
# Secrets
# ============================================================

def load_secret():

    secret_env = os.environ.get("APP_SECRET_JSON")

    if not secret_env:
        raise ExportError(
            "APP_SECRET_JSON is not set"
        )

    try:
        return json.loads(secret_env)

    except json.JSONDecodeError:
        raise ExportError(
            "APP_SECRET_JSON is invalid JSON"
        ) from None


def resolve_config(config):

    meta_conf = config.get(
        "meta",
        {}
    )

    sheets_conf = config.get(
        "sheets",
        {}
    )

    worksheet_name = (
        sheets_conf.get(
            "worksheet_name"
        )
        or DEFAULT_WORKSHEET_NAME
    )

    google_service_account = (
        config.get(
            "gcp_service_account"
        )
    )

    google_service_account = (
        normalize_google_service_account(
            google_service_account
        )
    )

    return {
        "meta": {
            "token":
                meta_conf.get(
                    "token"
                ),

            "account_id":
                meta_conf.get(
                    "account_id"
                ),
        },

        "sheet": {
            "spreadsheet_id":
                sheets_conf.get(
                    "spreadsheet_id"
                ),

            "worksheet_name":
                worksheet_name,

            "google_service_account":
                google_service_account,
        },
    }


def validate_config(resolved):

    required = {
        "meta.token":
            resolved["meta"]["token"],

        "meta.account_id":
            resolved["meta"]["account_id"],

        "sheets.spreadsheet_id":
            resolved["sheet"]["spreadsheet_id"],

        "gcp_service_account":
            resolved["sheet"]["google_service_account"],
    }

    missing = [
        key
        for key, value in required.items()
        if not value
    ]

    if missing:

        raise ExportError(
            "Missing required config keys: "
            + ", ".join(missing)
        )


    token = resolved["meta"]["token"]
    if not isinstance(token, str) or not token.strip() or token != token.strip():
        raise ExportError("Invalid Meta access token.")


def normalize_google_service_account(creds):

    if not creds:
        return None

    fixed = dict(
        creds
    )

    private_key = fixed.get(
        "private_key",
        ""
    )

    if private_key:

        fixed["private_key"] = (
            private_key.replace(
                "\\n",
                "\n"
            )
        )

    return fixed


# ============================================================
# 日付処理
# ============================================================

def add_months(
    base_date,
    months
):

    month = (
        base_date.month
        - 1
        + months
    )

    year = (
        base_date.year
        + month // 12
    )

    month = (
        month % 12
        + 1
    )

    return date(
        year,
        month,
        1
    )


def get_target_date_ranges():

    today_jst = datetime.now(
        JST
    ).date()

    yesterday = (
        today_jst
        - timedelta(days=1)
    )

    this_month_start = date(
        today_jst.year,
        today_jst.month,
        1
    )

    last_month_end = (
        this_month_start
        - timedelta(days=1)
    )

    last_month_start = date(
        last_month_end.year,
        last_month_end.month,
        1
    )

    # --------------------------------------------------------
    # 当月に取得可能日があるか
    # --------------------------------------------------------

    if yesterday >= this_month_start:

        latest_month = (
            this_month_start
        )

    else:

        latest_month = (
            last_month_start
        )

    # --------------------------------------------------------
    # month：過去12ヶ月
    # --------------------------------------------------------

    month_start_12 = add_months(
        latest_month,
        -11
    )

    monthly_12_ranges = (
        build_month_ranges(
            month_start_12,
            latest_month,
            yesterday
        )
    )

    # --------------------------------------------------------
    # month以外：過去2ヶ月
    # --------------------------------------------------------

    month_start_2 = add_months(
        latest_month,
        -1
    )

    monthly_2_ranges = (
        build_month_ranges(
            month_start_2,
            latest_month,
            yesterday
        )
    )

    # --------------------------------------------------------
    # day：過去2ヶ月
    # --------------------------------------------------------

    daily_since = (
        month_start_2
    )

    daily_until = (
        yesterday
    )

    return (
        monthly_12_ranges,
        monthly_2_ranges,
        daily_since,
        daily_until,
    )


def build_month_ranges(
    first_month,
    last_month,
    yesterday
):

    ranges = []

    current = (
        first_month
    )

    while current <= last_month:

        next_month = add_months(
            current,
            1
        )

        month_end = (
            next_month
            - timedelta(days=1)
        )

        until = min(
            month_end,
            yesterday
        )

        if until >= current:

            ranges.append({
                "month":
                    current,

                "since":
                    current,

                "until":
                    until,
            })

        current = (
            next_month
        )

    return ranges


def date_to_sheet_serial(
    value
):

    if not value:
        return ""

    return (
        value
        - GOOGLE_SHEETS_EPOCH
    ).days


# ============================================================
# Meta ID
# ============================================================

def normalize_meta_act_id(
    raw_act_id
):

    cleaned = (
        str(raw_act_id)
        .replace(
            "act=",
            ""
        )
        .replace(
            "act_",
            ""
        )
        .replace(
            "act",
            ""
        )
        .strip()
    )

    if not re.fullmatch(r"[0-9]+", cleaned):
        raise ExportError("Invalid Meta account ID.")
    return f"act_{cleaned}"


# ============================================================
# Metaデータ取得
# ============================================================

def fetch_meta_rows(
    act_id,
    token,
    monthly_12_ranges,
    monthly_2_ranges,
    daily_since,
    daily_until,
):

    act_id = normalize_meta_act_id(
        act_id
    )

    rows = []

    # ========================================================
    # 1. month
    # 月別・アカウント全体
    # 過去12ヶ月
    # ========================================================

    for month_range in monthly_12_ranges:

        data = fetch_meta_insights(
            act_id=act_id,
            token=token,
            since=month_range["since"],
            until=month_range["until"],
            level="account",
            time_increment="all_days",
            fields=[
                "actions",
            ],
        )

        for item in data:

            metrics = (
                extract_lead_metrics(
                    item
                )
            )

            rows.append(
                make_output_row(
                    scope="month",

                    month=month_range[
                        "month"
                    ],

                    cv_click_1d=metrics[
                        "cv_click_1d"
                    ],

                    cv_click_7d=metrics[
                        "cv_click_7d"
                    ],
                )
            )

    # ========================================================
    # 2. day
    # 日別・アカウント全体
    # 過去2ヶ月
    # ========================================================

    data = fetch_meta_insights(
        act_id=act_id,
        token=token,
        since=daily_since,
        until=daily_until,
        level="account",
        time_increment="1",
        fields=[
            "actions",
        ],
    )

    for item in data:

        day_str = item.get(
            "date_start"
        )

        day_value = (
            datetime.strptime(
                day_str,
                "%Y-%m-%d"
            ).date()
        )

        metrics = (
            extract_lead_metrics(
                item
            )
        )

        rows.append(
            make_output_row(
                scope="day",

                day=day_value,

                cv_click_1d=metrics[
                    "cv_click_1d"
                ],

                cv_click_7d=metrics[
                    "cv_click_7d"
                ],
            )
        )

    # ========================================================
    # 3～5.
    # ad / gen_age / pf
    # 過去2ヶ月
    # ========================================================

    for month_range in monthly_2_ranges:

        # ----------------------------------------------------
        # ad
        # ----------------------------------------------------

        ad_data = fetch_meta_insights(
            act_id=act_id,
            token=token,
            since=month_range["since"],
            until=month_range["until"],
            level="ad",
            time_increment="all_days",
            fields=[
                "campaign_name",
                "adset_name",
                "ad_name",
                "actions",
            ],
        )

        for item in ad_data:

            metrics = (
                extract_lead_metrics(
                    item
                )
            )

            rows.append(
                make_output_row(
                    scope="ad",

                    month=month_range[
                        "month"
                    ],

                    campaign_name=item.get(
                        "campaign_name",
                        ""
                    ),

                    adset_name=item.get(
                        "adset_name",
                        ""
                    ),

                    ad_name=item.get(
                        "ad_name",
                        ""
                    ),

                    cv_click_1d=metrics[
                        "cv_click_1d"
                    ],

                    cv_click_7d=metrics[
                        "cv_click_7d"
                    ],
                )
            )

        # ----------------------------------------------------
        # gen_age
        # 性別 × 年齢
        # ----------------------------------------------------

        gen_age_data = fetch_meta_insights(
            act_id=act_id,
            token=token,
            since=month_range["since"],
            until=month_range["until"],
            level="account",
            time_increment="all_days",
            fields=[
                "actions",
            ],
            breakdowns=[
                "gender",
                "age",
            ],
        )

        for item in gen_age_data:

            metrics = (
                extract_lead_metrics(
                    item
                )
            )

            rows.append(
                make_output_row(
                    scope="gen_age",

                    month=month_range[
                        "month"
                    ],

                    gender=item.get(
                        "gender",
                        ""
                    ),

                    age=item.get(
                        "age",
                        ""
                    ),

                    cv_click_1d=metrics[
                        "cv_click_1d"
                    ],

                    cv_click_7d=metrics[
                        "cv_click_7d"
                    ],
                )
            )

        # ----------------------------------------------------
        # pf
        # プラットフォーム別
        # ----------------------------------------------------

        platform_data = fetch_meta_insights(
            act_id=act_id,
            token=token,
            since=month_range["since"],
            until=month_range["until"],
            level="account",
            time_increment="all_days",
            fields=[
                "actions",
            ],
            breakdowns=[
                "publisher_platform",
            ],
        )

        for item in platform_data:

            metrics = (
                extract_lead_metrics(
                    item
                )
            )

            rows.append(
                make_output_row(
                    scope="pf",

                    month=month_range[
                        "month"
                    ],

                    platform=item.get(
                        "publisher_platform",
                        ""
                    ),

                    cv_click_1d=metrics[
                        "cv_click_1d"
                    ],

                    cv_click_7d=metrics[
                        "cv_click_7d"
                    ],
                )
            )

    return rows


# ============================================================
# Meta Insights API
# ============================================================

def meta_page(url, headers, params):
    try:
        response = requests.get(url, headers=headers, params=params, timeout=120, allow_redirects=False)
    except requests.RequestException:
        raise ExportError("Meta API connection failed.") from None
    status = response.status_code
    try:
        data = response.json()
    except ValueError:
        raise ExportError(f"Meta API invalid JSON (HTTP {status}).") from None
    has_error = isinstance(data, dict) and "error" in data
    if not 200 <= status < 300 or has_error:
        error = data.get("error") if has_error else None
        details = ""
        if isinstance(error, dict):
            for key in ("code", "error_subcode"):
                if type(error.get(key)) is int:
                    details += f", {key}={error[key]}"
        raise ExportError(f"Meta API failed (HTTP {status}{details}).")
    return data


def next_cursor(data, endpoint):
    paging = data.get("paging", {})
    if not isinstance(paging, dict):
        raise ExportError("Meta returned invalid paging.")
    if "next" not in paging:
        return None
    link = paging["next"]
    if not isinstance(link, str):
        raise ExportError("Meta returned an invalid next page.")
    parts = urlsplit(link)
    if (parts.scheme != "https" or parts.netloc != "graph.facebook.com"
            or parts.path != urlsplit(endpoint).path or parts.fragment):
        raise ExportError("Meta returned an unexpected next page.")
    query = parse_qs(parts.query, keep_blank_values=True)
    after = query.get("after", [])
    if len(after) != 1 or not after[0] or any(k in query for k in ("before", "offset", "__paging_token")):
        raise ExportError("Meta returned unsupported pagination.")
    cursors = paging.get("cursors", {})
    if not isinstance(cursors, dict) or cursors.get("after", after[0]) != after[0]:
        raise ExportError("Meta returned inconsistent cursors.")
    return after[0]


def fetch_meta_insights(
    act_id, token, since, until, level, time_increment, fields, breakdowns=None,
):
    url = f"https://graph.facebook.com/{META_API_VERSION}/{act_id}/insights"
    params = {
        "level": level,
        "time_range": json.dumps({"since": since.isoformat(), "until": until.isoformat()}),
        "fields": ",".join(fields),
        "time_increment": time_increment,
        "action_attribution_windows": json.dumps(["1d_click", "7d_click"]),
        "use_account_attribution_setting": "false",
        "use_unified_attribution_setting": "false",
        "limit": 5000,
    }
    if breakdowns:
        params["breakdowns"] = ",".join(breakdowns)

    headers = {"Authorization": f"Bearer {token}"}
    rows = []
    seen = set()
    while True:
        data = meta_page(url, headers, params)
        if not isinstance(data, dict) or not isinstance(data.get("data"), list):
            raise ExportError("Meta returned an invalid data array.")
        for row in data["data"]:
            validate_meta_row(row, fields, breakdowns or (), since, until, time_increment == "1")
        rows.extend(data["data"])
        after = next_cursor(data, url)
        if after is None:
            return rows
        if token in after:
            raise ExportError("Meta returned credentials in the page cursor.")
        if after in seen:
            raise ExportError("Meta pagination repeated a cursor.")
        seen.add(after)
        params["after"] = after


def validate_meta_row(row, fields, breakdowns, since, until, daily):
    if not isinstance(row, dict):
        raise ExportError("Meta returned an invalid row.")
    for field in fields:
        if field.endswith("_name") and not isinstance(row.get(field), str):
            raise ExportError("Meta returned an invalid ad hierarchy name.")
    for key in breakdowns:
        if not isinstance(row.get(key), str) or not row[key].strip():
            raise ExportError("Meta omitted a requested breakdown.")
    if daily:
        try:
            raw = row["date_start"]
            day = date.fromisoformat(raw)
            if not since <= day <= until or day.isoformat() != raw:
                raise ValueError
        except (KeyError, TypeError, ValueError):
            raise ExportError("Meta returned an invalid report date.") from None


# ============================================================
# CV抽出
# ============================================================

def extract_lead_metrics(item):
    actions = item.get("actions", [])
    if not isinstance(actions, list):
        raise ExportError("Meta returned an invalid actions array.")
    lead = None
    for action in actions:
        if not isinstance(action, dict) or not isinstance(action.get("action_type"), str) or not action["action_type"]:
            raise ExportError("Meta returned an invalid action.")
        if action["action_type"] == TARGET_ACTION_TYPE:
            if lead is not None:
                raise ExportError("Meta returned duplicate website lead actions.")
            lead = action
    if lead is None:
        return {"cv_click_1d": 0, "cv_click_7d": 0}
    if "1d_click" not in lead and "7d_click" not in lead:
        raise ExportError("Meta returned a website lead without the requested attribution windows.")
    return {
        "cv_click_1d": to_number(lead.get("1d_click", 0)),
        "cv_click_7d": to_number(lead.get("7d_click", 0)),
    }


# ============================================================
# 出力行
# ============================================================

def make_output_row(
    scope,
    month=None,
    day=None,
    campaign_name="",
    adset_name="",
    ad_name="",
    gender="",
    age="",
    platform="",
    cv_click_1d=0,
    cv_click_7d=0,
):

    return [

        # A
        "meta",

        # B
        scope,

        # C
        date_to_sheet_serial(
            month
        ),

        # D
        date_to_sheet_serial(
            day
        ),

        # E
        campaign_name or "",

        # F
        adset_name or "",

        # G
        ad_name or "",

        # H
        gender or "",

        # I
        age or "",

        # J
        platform or "",

        # K
        to_number(
            cv_click_1d
        ),

        # L
        to_number(
            cv_click_7d
        ),
    ]


# ============================================================
# Google Sheets
# ============================================================

def connect_spreadsheet(
    sheet_id,
    google_creds_dict,
):

    try:

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]

        creds = (
            Credentials
            .from_service_account_info(
                google_creds_dict,
                scopes=scopes
            )
        )

        client = (
            gspread.authorize(
                creds
            )
        )

        spreadsheet = (
            client.open_by_key(
                sheet_id
            )
        )

        print(
            "Google Sheets connected successfully"
        )

        return spreadsheet

    except Exception:

        raise ExportError(
            "Google Sheets connection failed"
        ) from None


HEADERS = [
        "media",
        "scope",
        "month",
        "day",
        "campaign_name",
        "adset_name",
        "ad_name",
        "gender",
        "age",
        "platform",
        "cv_click_1d",
        "cv_click_7d",
]


LEASE_SECONDS = 3600


def sheet_state(book):
    return book.fetch_sheet_metadata(params={
        "fields": "sheets(properties(sheetId,title,sheetType,gridProperties(rowCount,columnCount))),"
                  "namedRanges(namedRangeId,name)"
    })


def range_names(state):
    return {item["name"]: item["namedRangeId"] for item in state.get("namedRanges", [])}


def delete_range(range_id):
    return {"deleteNamedRange": {"namedRangeId": range_id}}


def add_range(name, range_id, sheet_id):
    return {"addNamedRange": {"namedRange": {
        "name": name, "namedRangeId": range_id,
        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                  "startColumnIndex": 0, "endColumnIndex": 1},
    }}}


def lease_time(range_id):
    match = re.fullmatch(r"bike_([0-9]+)_[0-9a-f]{32}", range_id)
    if not match:
        raise ExportError("Unexpected Sheets lock. Check named ranges.")
    return int(match.group(1))


def acquire_lease(book, title):
    # All runners compete for one unique name. Each acquisition uses a fresh ID.
    key = hashlib.sha256(title.casefold().encode()).hexdigest()[:24]
    lock = "bike_lock_" + key
    state = sheet_state(book)
    existing = range_names(state).get(lock)
    now = int(time.time())
    if existing and now < lease_time(existing) + LEASE_SECONDS:
        return None
    token = f"bike_{now}_{uuid.uuid4().hex}"
    anchor = next(s["properties"]["sheetId"] for s in state["sheets"]
                  if s["properties"].get("sheetType", "GRID") == "GRID")
    lease = {"id": token, "name": lock, "receipt_prefix": "bike_done_" + key + "_",
             "anchor": anchor}
    ops = [delete_range(existing)] if existing else []
    ops.append(add_range(lock, token, anchor))
    try:
        book.batch_update({"requests": ops})
    except Exception:
        owner = range_names(sheet_state(book)).get(lock)
        if owner == token:
            return lease
        if owner and owner != existing:
            return None
        raise ExportError("Sheets lock acquisition failed.") from None
    return lease


def release_lease(book, lease):
    # IDs are never reused. A delayed cleanup cannot unlock the successor.
    try:
        if range_names(sheet_state(book)).get(lease["name"]) == lease["id"]:
            book.batch_update({"requests": [delete_range(lease["id"])]})
    except Exception:
        print("Sheets lock cleanup unconfirmed. Expiry permits recovery.", file=sys.stderr)


def previous_report_height(book, title, new_height, column_count):
    escaped = title.replace("'", "''")
    width = min(column_count, len(HEADERS))
    last_column = chr(ord("A") + width - 1)
    result = book.values_get(
        f"'{escaped}'!A:{last_column}", params={"valueRenderOption": "FORMULA"}
    )
    values = result.get("values", [])
    height = 0
    if width >= 2 and values and values[0] == HEADERS[:width]:
        height = 1
        for row in values[1:]:
            if len(row) < 2 or row[0] != "meta":
                break
            if row[1] not in {"month", "day", "ad", "gen_age", "pf"}:
                break
            height += 1
    if any(any(cell != "" for cell in row) for row in values[height:new_height]):
        raise ExportError("Sheets output overlaps existing non-report data.")
    return height


def write_to_sheet(book, title, rows, lease):
    state = sheet_state(book)
    names = range_names(state)
    if names.get(lease["name"]) != lease["id"]:
        raise ExportError("Sheets lock ownership changed.")
    # A missing owner ID invalidates this *whole* batch even after the preliminary read.
    receipt_id = "receipt_" + lease["id"]
    receipt = lease["receipt_prefix"] + lease["id"]
    ops = [{"updateNamedRange": {"namedRange": {"namedRangeId": lease["id"], "name": receipt},
                                 "fields": "name"}},
           delete_range(lease["id"]), add_range(receipt, receipt_id, lease["anchor"])]
    # Keep receipts long enough for a timed-out writer to confirm its own commit.
    cutoff = int(time.time()) - 86400
    for name, range_id in names.items():
        saved_token = range_id.removeprefix("receipt_")
        if name == lease["receipt_prefix"] + saved_token and lease_time(saved_token) < cutoff:
            ops.append(delete_range(range_id))
    props = next((s["properties"] for s in state["sheets"]
                  if s["properties"]["title"] == title), None)
    output = [HEADERS, *rows]
    if props is None:
        sheet_id = secrets.randbelow(2**31 - 1) + 1
        height = len(output)
        ops.append({"addSheet": {"properties": {
            "sheetId": sheet_id, "title": title,
            "gridProperties": {"rowCount": max(1000, len(rows) + 100), "columnCount": len(HEADERS)},
        }}})
    else:
        sheet_id = props["sheetId"]
        grid = props["gridProperties"]
        height = max(len(output), previous_report_height(book, title, len(output), grid["columnCount"]))
        for dimension, current, needed in (("ROWS", grid["rowCount"], max(height, 100)),
                                            ("COLUMNS", grid["columnCount"], len(HEADERS))):
            if needed > current:
                ops.append({"appendDimension": {"sheetId": sheet_id, "dimension": dimension,
                                                "length": needed - current}})
    # updateCells with an explicit range clears only surplus values in the old report.
    cell_rows = []
    for row in output:
        cells = []
        for value in row:
            field = "numberValue" if isinstance(value, (int, float)) else "stringValue"
            cells.append({"userEnteredValue": {field: value}})
        cell_rows.append({"values": cells})
    ops.append({"updateCells": {
        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": height,
                  "startColumnIndex": 0, "endColumnIndex": len(HEADERS)},
        "rows": cell_rows, "fields": "userEnteredValue",
    }})
    if rows:
        ops.append({"repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(output),
                      "startColumnIndex": 2, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "yyyy-mm-dd"}}},
            "fields": "userEnteredFormat.numberFormat",
        }})
    ops.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount",
    }})
    try:
        book.batch_update({"requests": ops})
    except Exception:
        try:
            confirmed = range_names(sheet_state(book)).get(receipt) == receipt_id
        except Exception:
            raise ExportError("Sheets commit could not be confirmed.") from None
        if not confirmed:
            raise ExportError("Sheets commit was not confirmed.") from None



# ============================================================
# 並び替え
# ============================================================

def sort_rows(
    rows
):

    scope_order = {
        "month":
            0,

        "day":
            1,

        "ad":
            2,

        "gen_age":
            3,

        "pf":
            4,
    }

    def sort_key(
        row
    ):

        (
            _media,
            scope,
            month,
            day,
            campaign_name,
            adset_name,
            ad_name,
            gender,
            age,
            platform,
            _click_1d,
            _click_7d,
        ) = row

        if day not in (
            "",
            None
        ):

            period = (
                day
            )

        else:

            period = (
                month
            )

        if not isinstance(
            period,
            (int, float)
        ):

            period = (
                0
            )

        return (
            scope_order.get(
                scope,
                999
            ),

            -period,

            campaign_name,
            adset_name,
            ad_name,
            gender,
            age,
            platform,
        )

    return sorted(
        rows,
        key=sort_key
    )


# ============================================================
# 数値処理
# ============================================================

def to_number(value):
    try:
        if isinstance(value, bool):
            raise ValueError
        num = float(value)
        if not math.isfinite(num) or num < 0:
            raise ValueError
    except (ValueError, TypeError, OverflowError):
        raise ExportError("Meta returned an invalid website lead count.") from None
    return int(num) if num.is_integer() else num


# ============================================================
# Run
# ============================================================

if __name__ == "__main__":
    logging.disable(logging.CRITICAL)
    try:
        main()
    except ExportError as exc:
        sys.exit(str(exc))
    except Exception:
        sys.exit("Export failed: unexpected configuration, service, or response error.")
