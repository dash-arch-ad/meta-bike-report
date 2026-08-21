import os
import json
import requests
import gspread

from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials


# ============================================================
# 基本設定
# ============================================================

META_API_VERSION = "v25.0"

JST = ZoneInfo("Asia/Tokyo")

DEFAULT_WORKSHEET_NAME = "gitreport"

# Webサイト上のLeadイベント
TARGET_ACTION_TYPE = "offsite_conversion.fb_pixel_lead"

# Google Sheetsの日付シリアル値の基準日
GOOGLE_SHEETS_EPOCH = date(1899, 12, 30)


# ============================================================
# main
# ============================================================

def main():

    print("=== Start Meta Lead Export ===")

    config = load_secret()

    mask_sensitive_values(config)

    resolved = resolve_config(config)

    validate_config(resolved)

    # --------------------------------------------------------
    # 対象期間
    # --------------------------------------------------------

    monthly_12_ranges, monthly_2_ranges, daily_since, daily_until = (
        get_target_date_ranges()
    )

    print("Month scope ranges:")
    for r in monthly_12_ranges:
        print(
            f"  {r['month']} : "
            f"{r['since']} - {r['until']}"
        )

    print("Other monthly scopes:")
    for r in monthly_2_ranges:
        print(
            f"  {r['month']} : "
            f"{r['since']} - {r['until']}"
        )

    print(
        f"Day scope range: "
        f"{daily_since} - {daily_until}"
    )

    # --------------------------------------------------------
    # Meta取得
    # --------------------------------------------------------

    rows = fetch_meta_rows(
        act_id=resolved["meta"]["account_id"],
        token=resolved["meta"]["token"],
        monthly_12_ranges=monthly_12_ranges,
        monthly_2_ranges=monthly_2_ranges,
        daily_since=daily_since,
        daily_until=daily_until,
    )

    print(f"Meta rows built: {len(rows)}")

    rows = sort_rows(rows)

    # --------------------------------------------------------
    # Google Sheets接続
    # --------------------------------------------------------

    spreadsheet = connect_spreadsheet(
        sheet_id=resolved["sheet"]["spreadsheet_id"],
        google_creds_dict=resolved["sheet"]["google_service_account"],
    )

    # --------------------------------------------------------
    # 書き込み
    # --------------------------------------------------------

    write_to_sheet(
        spreadsheet=spreadsheet,
        sheet_name=resolved["sheet"]["worksheet_name"],
        rows=rows,
    )

    print(f"Total rows written: {len(rows)}")

    print("=== Completed ===")


# ============================================================
# Secrets
# ============================================================

def load_secret():

    secret_env = os.environ.get("APP_SECRET_JSON")

    if not secret_env:
        raise RuntimeError(
            "APP_SECRET_JSON is not set"
        )

    try:
        return json.loads(secret_env)

    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"APP_SECRET_JSON is invalid JSON: {e}"
        ) from e


def mask_sensitive_values(config):

    values = []

    meta = config.get("meta", {})

    for value in [
        meta.get("token"),
        meta.get("account_id"),
    ]:

        if value is None:
            continue

        value = str(value).strip()

        if not value:
            continue

        if "\n" in value:
            continue

        values.append(value)

    for value in sorted(set(values)):
        print(f"::add-mask::{value}")


def resolve_config(config):

    meta_conf = config.get("meta", {})
    sheets_conf = config.get("sheets", {})

    worksheet_name = (
        sheets_conf.get("worksheet_name")
        or DEFAULT_WORKSHEET_NAME
    )

    google_service_account = (
        config.get("gcp_service_account")
    )

    google_service_account = (
        normalize_google_service_account(
            google_service_account
        )
    )

    return {
        "meta": {
            "token": meta_conf.get("token"),
            "account_id": meta_conf.get("account_id"),
        },

        "sheet": {
            "spreadsheet_id":
                sheets_conf.get("spreadsheet_id"),

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
        raise RuntimeError(
            "Missing required config keys: "
            + ", ".join(missing)
        )


def normalize_google_service_account(creds):

    if not creds:
        return None

    fixed = dict(creds)

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

def add_months(base_date, months):

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
    """
    例：
    2026/08/21 実行

    month
        2025/09 ～ 2026/08
        過去12ヶ月

    day
        2026/07/01 ～ 2026/08/20

    ad / gen_age / pf
        2026/07
        2026/08
    """

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
    # 当月に1日以上データがあるか
    # --------------------------------------------------------

    if yesterday >= this_month_start:
        latest_month = this_month_start
    else:
        latest_month = last_month_start

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
    # その他：過去2ヶ月
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
    # day：前月1日～前日
    # --------------------------------------------------------

    daily_since = month_start_2
    daily_until = yesterday

    return (
        monthly_12_ranges,
        monthly_2_ranges,
        daily_since,
        daily_until,
    )


def build_month_ranges(
    first_month,
    last_month,
    yesterday,
):

    ranges = []

    current = first_month

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
                "month": current,
                "since": current,
                "until": until,
            })

        current = next_month

    return ranges


def date_to_sheet_serial(value):

    if not value:
        return ""

    return (
        value
        - GOOGLE_SHEETS_EPOCH
    ).days


# ============================================================
# Meta ID
# ============================================================

def normalize_meta_act_id(raw_act_id):

    cleaned = (
        str(raw_act_id)
        .replace("act=", "")
        .replace("act_", "")
        .replace("act", "")
        .strip()
    )

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

            metrics = extract_lead_metrics(
                item
            )

            rows.append(
                make_output_row(
                    scope="month",
                    month=month_range["month"],
                    cv_view_1d=metrics[
                        "cv_view_1d"
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

        if not day_str:
            continue

        day_value = datetime.strptime(
            day_str,
            "%Y-%m-%d"
        ).date()

        metrics = extract_lead_metrics(
            item
        )

        rows.append(
            make_output_row(
                scope="day",
                day=day_value,
                cv_view_1d=metrics[
                    "cv_view_1d"
                ],
                cv_click_7d=metrics[
                    "cv_click_7d"
                ],
            )
        )

    # ========================================================
    # 3～5. 月別系
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

            metrics = extract_lead_metrics(
                item
            )

            rows.append(
                make_output_row(
                    scope="ad",
                    month=month_range["month"],
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
                    cv_view_1d=metrics[
                        "cv_view_1d"
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

            metrics = extract_lead_metrics(
                item
            )

            rows.append(
                make_output_row(
                    scope="gen_age",
                    month=month_range["month"],
                    gender=item.get(
                        "gender",
                        ""
                    ),
                    age=item.get(
                        "age",
                        ""
                    ),
                    cv_view_1d=metrics[
                        "cv_view_1d"
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

            metrics = extract_lead_metrics(
                item
            )

            rows.append(
                make_output_row(
                    scope="pf",
                    month=month_range["month"],
                    platform=item.get(
                        "publisher_platform",
                        ""
                    ),
                    cv_view_1d=metrics[
                        "cv_view_1d"
                    ],
                    cv_click_7d=metrics[
                        "cv_click_7d"
                    ],
                )
            )

    return rows


# ============================================================
# Insights API
# ============================================================

def fetch_meta_insights(
    act_id,
    token,
    since,
    until,
    level,
    time_increment,
    fields,
    breakdowns=None,
):

    url = (
        f"https://graph.facebook.com/"
        f"{META_API_VERSION}/"
        f"{act_id}/insights"
    )

    params = {
        "access_token": token,

        "level": level,

        "time_range": json.dumps({
            "since":
                since.strftime(
                    "%Y-%m-%d"
                ),

            "until":
                until.strftime(
                    "%Y-%m-%d"
                ),
        }),

        "fields":
            ",".join(fields),

        "time_increment":
            time_increment,

        # 指定したアトリビューション期間
        "action_attribution_windows":
            json.dumps([
                "1d_view",
                "7d_click",
            ]),

        # アカウント側の既定値で上書きさせず、
        # 上記windowを使用
        "use_account_attribution_setting":
            "false",

        "use_unified_attribution_setting":
            "false",

        "limit": 5000,
    }

    if breakdowns:

        params["breakdowns"] = (
            ",".join(breakdowns)
        )

    all_rows = []

    while True:

        response = requests.get(
            url,
            params=params,
            timeout=120
        )

        try:
            response.raise_for_status()

        except requests.HTTPError as e:

            raise RuntimeError(
                "Meta API request failed. "
                f"status={response.status_code}, "
                f"body={truncate_text(response.text)}"
            ) from e

        payload = response.json()

        if "error" in payload:

            raise RuntimeError(
                "Meta API error: "
                + json.dumps(
                    payload["error"],
                    ensure_ascii=False
                )
            )

        batch = payload.get(
            "data",
            []
        )

        all_rows.extend(
            batch
        )

        next_url = (
            payload
            .get("paging", {})
            .get("next")
        )

        if not next_url:
            break

        url = next_url
        params = None

    return all_rows


# ============================================================
# CV抽出
# ============================================================

def extract_lead_metrics(item):
    """
    actions の中から

    offsite_conversion.fb_pixel_lead

    の

    1d_view
    7d_click

    を取り出す。
    """

    result = {
        "cv_view_1d": 0,
        "cv_click_7d": 0,
    }

    actions = item.get(
        "actions",
        []
    )

    if not isinstance(actions, list):
        return result

    for action in actions:

        if (
            action.get("action_type")
            != TARGET_ACTION_TYPE
        ):
            continue

        result["cv_view_1d"] = (
            to_number(
                action.get(
                    "1d_view",
                    0
                )
            )
        )

        result["cv_click_7d"] = (
            to_number(
                action.get(
                    "7d_click",
                    0
                )
            )
        )

        break

    return result


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
    cv_view_1d=0,
    cv_click_7d=0,
):

    return [
        "meta",
        scope,

        date_to_sheet_serial(
            month
        ),

        date_to_sheet_serial(
            day
        ),

        campaign_name or "",
        adset_name or "",
        ad_name or "",
        gender or "",
        age or "",
        platform or "",

        to_number(
            cv_view_1d
        ),

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

        client = gspread.authorize(
            creds
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

    except Exception as e:

        raise RuntimeError(
            "Google Sheets connection error: "
            f"{repr(e)}"
        ) from e


def write_to_sheet(
    spreadsheet,
    sheet_name,
    rows,
):

    header = [[
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
        "cv_view_1d",
        "cv_click_7d",
    ]]

    try:

        try:
            worksheet = (
                spreadsheet.worksheet(
                    sheet_name
                )
            )

        except gspread.WorksheetNotFound:

            worksheet = (
                spreadsheet.add_worksheet(
                    title=sheet_name,
                    rows=max(
                        1000,
                        len(rows) + 100
                    ),
                    cols=12,
                )
            )

        # 必要サイズ確保
        required_rows = max(
            len(rows) + 1,
            100
        )

        if worksheet.row_count < required_rows:

            worksheet.resize(
                rows=required_rows
            )

        if worksheet.col_count < 12:

            worksheet.resize(
                cols=12
            )

        # 既存データ削除
        worksheet.clear()

        output = (
            header
            + rows
        )

        # RAWで書くことにより
        # C/D列の日付シリアル値を数値として保持
        worksheet.update(
            range_name="A1",
            values=output,
            value_input_option="RAW",
        )

        # C/D列を日付表示形式に設定
        format_date_columns(
            spreadsheet=spreadsheet,
            worksheet=worksheet,
            row_count=len(output),
        )

        # 1行目固定
        worksheet.freeze(
            rows=1
        )

        print(
            f"Write success: "
            f"{sheet_name} "
            f"({len(rows)} rows)"
        )

    except Exception as e:

        raise RuntimeError(
            f"Write error ({sheet_name}): "
            f"{repr(e)}"
        ) from e


def format_date_columns(
    spreadsheet,
    worksheet,
    row_count,
):

    if row_count <= 1:
        return

    sheet_id = worksheet.id

    request = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId":
                            sheet_id,

                        "startRowIndex":
                            1,

                        "endRowIndex":
                            row_count,

                        # C列
                        "startColumnIndex":
                            2,

                        # D列まで
                        "endColumnIndex":
                            4,
                    },

                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "DATE",
                                "pattern":
                                    "yyyy-mm-dd",
                            }
                        }
                    },

                    "fields":
                        "userEnteredFormat.numberFormat",
                }
            }
        ]
    }

    spreadsheet.batch_update(
        request
    )


# ============================================================
# 並び替え
# ============================================================

def sort_rows(rows):

    scope_order = {
        "month": 0,
        "day": 1,
        "ad": 2,
        "gen_age": 3,
        "pf": 4,
    }

    def sort_key(row):

        (
            media,
            scope,
            month,
            day,
            campaign_name,
            adset_name,
            ad_name,
            gender,
            age,
            platform,
            _view,
            _click,
        ) = row

        period = (
            day
            if day not in ("", None)
            else month
        )

        period = (
            period
            if isinstance(
                period,
                (int, float)
            )
            else 0
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

    if value in (
        None,
        ""
    ):
        return 0

    try:

        number = float(
            value
        )

        if number.is_integer():
            return int(
                number
            )

        return number

    except (
        TypeError,
        ValueError
    ):
        return 0


def truncate_text(
    value,
    limit=1000,
):

    value = str(
        value
    )

    if len(value) <= limit:
        return value

    return (
        value[:limit]
        + "...(truncated)"
    )


# ============================================================
# Run
# ============================================================

if __name__ == "__main__":
    main()
