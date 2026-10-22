from __future__ import annotations

import copy
import datetime as dt
import io
import re
import zipfile
from collections import Counter
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter


# Source fields used in the current One Stop export
SOURCE_QUESTION_FIELDS = {
    "location_description": [
        "Please describe the location of the store (i.e. names of the stores on either side,  or local landmarks):",
        "Please describe the location of the store (i.e. names of the stores on either side, or local landmarks):",
    ],
    "alcohol_type": [
        "What type of alcohol did you try to purchase?",
        "What type of alcohol did you purchase?",
    ],
    "alcohol_details": [
        "Please give details of the alcohol that you purchased:",
        "Please give details of the alcohol purchased (brand and size):",
    ],
    "supervisor": [
        "Was a supervisor called at any point during the transaction?",
        "Was a supervisor called at any time during the transaction?",
    ],
    "asked_id": [
        "Did the staff member who served you ask for ID?",
        "Did the person (or their supervisor) ask you for ID during the transaction?",
    ],
    "eye_contact": [
        "Did the staff member who served you make eye contact with you during the transaction?",
        "Did the person who served you make eye contact with you during the transaction?",
    ],
    "when_eye_contact": [
        "When was eye contact first made?",
        "If eye contact was made, when did the person who served you FIRST make eye contact?",
    ],
    "assess_age": [
        "Did the staff member who served you look at you long enough to assess your age?",
        "Did the server look at you long enough to make an assessment of your age?",
    ],
    "staff_gender": [
        "What was the gender of the staff member who served you?",
        "What was the gender of the server?",
    ],
    "staff_age": [
        "What was the approximate age of the staff member?",
        "What was the approximate age of the server?",
    ],
    "staff_name": [
        "What was the name of the staff member who served you?",
        "What was the name of the server (from the receipt or name badge):",
    ],
    "staff_friendly": [
        "Were the staff members who interacted with you friendly and polite?",
        "Were the members of staff who interacted with you friendly and polite?",
    ],
    "unfriendly_reason": [
        "Please describe the impolite/unfriendly interaction with staff members:",
        "If not, then give a reason:",
    ],
    "receipt_address": [
        "From the top of the receipt, please enter the postcode or any available address details:",
        "From the top of the receipt, please enter the post code or any available address details:",
    ],
    "receipt_store_code": [
        "From the bottom left of the receipt, please enter the Store code:",
        "From the top right hand side of the receipt, please enter the T CODE:",
    ],
    "think25": [
        "Did you see 'Think 25' notices in the store?",
        "Did you see \"Think 25\" notices in the store?",
    ],
    "underage_evidence": [
        "Did you see any evidence of anyone buying alcohol for underage people?",
    ],
    "underage_details": [
        "Please detail the evidence of anyone buying alcohol for underage people",
        "Please detail the evidence of anyone buying alcohol for underage people:",
    ],
    "unusual": [
        "Did anything unusual occur on the audit, or do you need to clarify any detail of your report:",
        "Please use this space to explain anything unusual about your visit or to clarify any detail of your report:",
        "Please use this space to explain anything unusual about your visit or to clarify any detail of your report:",
    ],
    "confirm_id": [
        "Please confirm below whether or not you were asked for ID:",
        "Please confirm whether or not you were asked for ID:",
        "Please confirm whether or not you were asked for ID, and if so, at what point during the transaction ID was checked:",
    ],
    "ecig_type": [
        "What type of E-cigarette product did you purchase/attempt to purchase?",
    ],
    "ecig_details": [
        "Please give details of the E-Cig that you purchased:",
        "Please give details of the e-cig product that you purchased:",
    ],
    "cigarette_details": [
        "Please give details of the cigarettes that you purchased:",
    ],
}

# Target worksheet column numbers, preserving the template's existing column positions.
TARGET_COLUMNS = {
    "order": 1, "client": 2, "visit": 3, "site": 4, "deadline": 5,
    "responsibility": 6, "premises": 7, "address1": 8, "address2": 9,
    "address3": 10, "city": 11, "postcode": 12, "submitted": 13,
    "approved": 14, "category": 15, "visit_date": 16, "visit_time": 17,
    "am_pm": 18, "result": 19, "result2": 20, "abort_reason": 21,
    "epos": 22, "category_duplicate": 26, "visitor_gender": 27,
    "location_description": 28, "product_type": 29, "product_details": 30,
    "supervisor": 31, "asked_id": 32, "eye_contact": 33,
    "when_eye_contact": 34, "assess_age": 35, "staff_gender": 36,
    "staff_age": 37, "staff_name": 38, "staff_friendly": 39,
    "unfriendly_reason": 40, "receipt_address": 41, "receipt_store_code": 42,
    "think25": 43, "underage_evidence": 44, "underage_details": 45,
    "unusual": 46, "confirm_id": 47,
}


def _normalise_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalise_category(value):
    """Return the Master Schedule category equivalent for an audit product."""
    s = _normalise_text(value).casefold().replace("-", "").replace(" ", "").replace("_", "")
    if s in {"alcohol", "alcoholic"}:
        return "Alcohol"
    if s in {"ecig", "ecigs", "ecigarette", "ecigarettes", "vape", "vapes"}:
        return "E-Cigs"
    if s in {"cigarette", "cigarettes", "cig", "cigs", "tobacco"}:
        return "Tobacco"
    return None


def _find_column(df: pd.DataFrame, candidates, required=False):
    lookup = {str(c).strip().casefold(): c for c in df.columns}
    for candidate in candidates:
        if candidate.casefold() in lookup:
            return lookup[candidate.casefold()]
    if required:
        raise ValueError(f"Required column not found. Expected one of: {', '.join(candidates)}")
    return None


def _read_master_schedule(file_bytes: bytes):
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None, dtype=object)
    header_idx = None
    for idx in range(min(6, len(raw))):
        row = [_normalise_text(v).casefold() for v in raw.iloc[idx].tolist()]
        if "ns id" in row and "epos code" in row:
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError("Could not find the 'NS ID' and 'EPoS Code' headers in the Master Schedule.")
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=header_idx, dtype=object)
    df.columns = [str(c).strip() if not pd.isna(c) else "" for c in df.columns]
    id_col = _find_column(df, ["NS ID", "SITEID", "Site ID"], required=True)
    epos_col = _find_column(df, ["EPoS Code", "EPOS Code", "Site Code"])
    period_cols = {str(c).strip().casefold(): c for c in df.columns}
    # Normalise identifiers and remove empty/duplicate IDs.
    df = df.copy()
    df["_site_id"] = df[id_col].map(_normalise_text)
    df = df[df["_site_id"].str.match(r"^SITE", case=False, na=False)]
    df = df.drop_duplicates("_site_id", keep="first")
    schedule = {}
    for _, row in df.iterrows():
        sid = row["_site_id"].upper()
        p_assignments = {}
        for p in range(1, 13):
            col = period_cols.get(f"period {p}".casefold())
            p_assignments[p] = normalise_category(row[col]) if col is not None else None
        epos = row[epos_col] if epos_col is not None else None
        if pd.isna(epos):
            epos = None
        schedule[sid] = {"epos": epos, "assignments": p_assignments}
    return schedule


def _read_period_calendar(file_bytes: bytes):
    cal = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, dtype=object)
    cal.columns = [str(c).strip() if not pd.isna(c) else "" for c in cal.columns]
    period_col = _find_column(cal, ["Period"], required=True)
    week_col = _find_column(cal, ["Week"], required=True)
    year_cols = [c for c in cal.columns if re.fullmatch(r"\d{2}-\d{2}", str(c).strip())]
    if not year_cols:
        raise ValueError("Could not find financial-year date columns such as '26-27' in the period calendar.")
    cal = cal.copy()
    cal["_period"] = pd.to_numeric(cal[period_col], errors="coerce").ffill()
    cal["_week"] = pd.to_numeric(cal[week_col], errors="coerce")
    return cal, year_cols


def period_dates(calendar_bytes: bytes, financial_year: str, period: int):
    cal, year_cols = _read_period_calendar(calendar_bytes)
    if financial_year not in year_cols:
        raise ValueError(f"Financial year '{financial_year}' is not present in the period calendar.")
    date_col = financial_year
    rows = cal[(cal["_period"] == int(period)) & cal[date_col].notna()].copy()
    if rows.empty:
        raise ValueError(f"No calendar dates found for Period {period}, financial year {financial_year}.")
    dates = pd.to_datetime(rows[date_col], errors="coerce").dropna().sort_values()
    if dates.empty:
        raise ValueError(f"No valid dates found for Period {period}, financial year {financial_year}.")
    # Tesco calendar lists Monday week-starts. One Stop starts Sunday and ends Saturday.
    start_date = (dates.iloc[0] - pd.Timedelta(days=1)).date()
    end_date = (dates.iloc[-1] + pd.Timedelta(days=5)).date()
    return start_date, end_date


def _source_value(row, candidates):
    for candidate in candidates:
        if candidate in row.index:
            value = row[candidate]
            if not pd.isna(value):
                return value
    # Case-insensitive fallback for question labels with punctuation/spacing differences.
    lookup = {str(c).strip().casefold(): c for c in row.index}
    for candidate in candidates:
        col = lookup.get(candidate.strip().casefold())
        if col is not None and not pd.isna(row[col]):
            return row[col]
    return None


def _as_excel_date(value):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return None
    d = pd.to_datetime(value, dayfirst=True, errors="coerce")
    if pd.isna(d):
        return None
    return d.to_pydatetime().replace(tzinfo=None)


def _as_excel_time(value):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return None
    if isinstance(value, dt.time):
        return value.replace(tzinfo=None)
    if isinstance(value, dt.datetime):
        return value.time().replace(tzinfo=None)
    parsed = pd.to_datetime(str(value), errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.time().replace(tzinfo=None)


def _is_aborted(row):
    # Any explicit abort indicator excludes the audit; ordinary FAIL outcomes remain included.
    for field in ["primary_result", "secondary_result", "tertiary_result", "status"]:
        val = _normalise_text(row.get(field, "")).casefold()
        if "abort" in val:
            return True
    success = _normalise_text(row.get("Were you able to successfully conduct this audit?", "")).casefold()
    if success in {"no", "no.", "no - aborted", "no, i was unable to conduct this audit", "no, i was unable to complete this audit"}:
        return True
    detail = _normalise_text(row.get("Please detail why you were unable to conduct this audit:", ""))
    if success and success.startswith("no") and detail:
        return True
    return False


def filter_audits(csv_bytes: bytes, master_bytes: bytes, calendar_bytes: bytes, financial_year: str, period: int):
    try:
        df = pd.read_csv(io.BytesIO(csv_bytes), dtype=object, keep_default_na=True)
    except UnicodeDecodeError:
        df = pd.read_csv(io.BytesIO(csv_bytes), dtype=object, keep_default_na=True, encoding="latin-1")
    df.columns = [str(c).strip() for c in df.columns]
    if "site_internal_id" not in df.columns:
        raise ValueError("The audit CSV must contain 'site_internal_id' so it can be matched to the Master Schedule.")
    if "item_to_order" not in df.columns:
        raise ValueError("The audit CSV must contain 'item_to_order' to identify Alcohol, E-Cig and Cigarettes.")
    if "date_of_visit" not in df.columns:
        raise ValueError("The audit CSV must contain 'date_of_visit' to assign audits to the selected period.")
    if "internal_id" not in df.columns:
        raise ValueError("The audit CSV must contain 'internal_id' for duplicate protection in YTD.")
    schedule = _read_master_schedule(master_bytes)
    start_date, end_date = period_dates(calendar_bytes, financial_year, int(period))

    df["_site_id"] = df["site_internal_id"].map(_normalise_text).str.upper()
    df["_category"] = df["item_to_order"].map(normalise_category)
    df["_visit_date"] = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.date
    df["_is_abort"] = df.apply(_is_aborted, axis=1)
    if "client_name" in df.columns:
        df["_client_ok"] = df["client_name"].map(_normalise_text).str.casefold().isin({"one stop", "onestop"})
    else:
        df["_client_ok"] = True
    df["_in_period"] = df["_visit_date"].map(lambda d: d is not None and not pd.isna(d) and start_date <= d <= end_date)
    df["_site_listed"] = df["_site_id"].isin(schedule.keys())
    df["_scheduled_category"] = df["_site_id"].map(
        lambda sid: schedule.get(sid, {}).get("assignments", {}).get(int(period))
    )
    df["_period_assignment_match"] = df["_category"].notna() & (df["_category"] == df["_scheduled_category"])

    # Apply the agreed rules: One Stop client, selected period, allowed product,
    # site listed, category assigned to that site for the selected period, and not aborted.
    mask = (
        df["_client_ok"]
        & df["_in_period"]
        & df["_category"].isin(["Alcohol", "E-Cigs", "Tobacco"])
        & df["_site_listed"]
        & df["_period_assignment_match"]
        & (~df["_is_abort"])
    )
    filtered = df.loc[mask].copy()
    # Keep the original export category labels for report readability.
    filtered = filtered.sort_values(
        by=["_visit_date", "site_name", "item_to_order"] if "site_name" in filtered.columns else ["_visit_date", "item_to_order"],
        kind="stable",
    )

    summary = {
        "source_rows": int(len(df)),
        "one_stop_rows": int(df["_client_ok"].sum()),
        "period_rows": int((df["_client_ok"] & df["_in_period"]).sum()),
        "allowed_category_rows": int((df["_client_ok"] & df["_in_period"] & df["_category"].isin(["Alcohol", "E-Cigs", "Tobacco"])).sum()),
        "listed_site_rows": int((df["_client_ok"] & df["_in_period"] & df["_category"].isin(["Alcohol", "E-Cigs", "Tobacco"]) & df["_site_listed"]).sum()),
        "scheduled_category_rows": int((df["_client_ok"] & df["_in_period"] & df["_category"].isin(["Alcohol", "E-Cigs", "Tobacco"]) & df["_site_listed"] & df["_period_assignment_match"]).sum()),
        "aborted_source_rows": int(df["_is_abort"].sum()),
        "included_rows": int(len(filtered)),
        "start_date": start_date,
        "end_date": end_date,
        "period": int(period),
        "financial_year": financial_year,
        "category_counts": filtered["item_to_order"].value_counts(dropna=False).to_dict(),
        "schedule": schedule,
    }
    return df, filtered, summary


def _copy_row_style_and_formulas(ws, source_row: int, target_row: int, max_col: int):
    if source_row == target_row:
        return
    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy.copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.value is not None and isinstance(src.value, str) and src.value.startswith("="):
            try:
                dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
            except Exception:
                dst.value = src.value
        else:
            dst.value = None
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def _value_from_row(row, candidates):
    return _source_value(row, candidates)


def _make_output_row(row, schedule_entry, result_case="upper"):
    category_original = _normalise_text(row.get("item_to_order", ""))
    result = _normalise_text(row.get("primary_result", ""))
    if result_case == "upper":
        result = result.upper()
    else:
        result = result.lower()
    if result in {"PASS", "pass"}:
        result = "PASS" if result_case == "upper" else "pass"
    elif result in {"FAIL", "fail"}:
        result = "FAIL" if result_case == "upper" else "fail"

    visit_time = _as_excel_time(row.get("time_of_visit"))
    am_pm = visit_time.strftime("%p") if visit_time else None
    epos = schedule_entry.get("epos") if schedule_entry else None
    if isinstance(epos, float) and epos.is_integer():
        epos = int(epos)

    # Common question columns are retained in their current template positions.
    if category_original.casefold() == "alcohol":
        product_type = _value_from_row(row, SOURCE_QUESTION_FIELDS["alcohol_type"])
        product_details = _value_from_row(row, SOURCE_QUESTION_FIELDS["alcohol_details"])
    elif category_original.casefold() in {"e-cig", "e-cigs", "ecig", "e-cigarette"}:
        product_type = _value_from_row(row, SOURCE_QUESTION_FIELDS["ecig_type"])
        product_details = _value_from_row(row, SOURCE_QUESTION_FIELDS["ecig_details"])
    else:
        product_type = "Cigs"
        product_details = _value_from_row(row, SOURCE_QUESTION_FIELDS["cigarette_details"])

    values = {
        1: _normalise_text(row.get("order_internal_id")),
        2: _normalise_text(row.get("client_name")) or "One Stop",
        3: _normalise_text(row.get("internal_id")),
        4: _normalise_text(row.get("site_internal_id")),
        5: _as_excel_date(row.get("deadline_date")),
        6: _normalise_text(row.get("responsibility")) or None,
        7: _normalise_text(row.get("site_name")) or None,
        8: _normalise_text(row.get("site_address_1")) or None,
        9: _normalise_text(row.get("site_address_2")) or None,
        10: _normalise_text(row.get("site_address_3")) or None,
        11: _normalise_text(row.get("site_address_4")) or None,
        12: _normalise_text(row.get("site_post_code")) or None,
        13: _as_excel_date(row.get("submitted_date")),
        14: _as_excel_date(row.get("approval_date")),
        15: category_original,
        16: _as_excel_date(row.get("date_of_visit")),
        17: visit_time,
        18: am_pm,
        19: result or None,
        20: _normalise_text(row.get("secondary_result")) or None,
        21: None,  # Aborted audits are excluded.
        22: epos,
        26: category_original,
        27: _normalise_text(row.get("auditor_gender")) or None,
        28: _value_from_row(row, SOURCE_QUESTION_FIELDS["location_description"]),
        29: product_type,
        30: product_details,
        31: _value_from_row(row, SOURCE_QUESTION_FIELDS["supervisor"]),
        32: _value_from_row(row, SOURCE_QUESTION_FIELDS["asked_id"]),
        33: _value_from_row(row, SOURCE_QUESTION_FIELDS["eye_contact"]),
        34: _value_from_row(row, SOURCE_QUESTION_FIELDS["when_eye_contact"]),
        35: _value_from_row(row, SOURCE_QUESTION_FIELDS["assess_age"]),
        36: _value_from_row(row, SOURCE_QUESTION_FIELDS["staff_gender"]),
        37: _value_from_row(row, SOURCE_QUESTION_FIELDS["staff_age"]),
        38: _value_from_row(row, SOURCE_QUESTION_FIELDS["staff_name"]),
        39: _value_from_row(row, SOURCE_QUESTION_FIELDS["staff_friendly"]),
        40: _value_from_row(row, SOURCE_QUESTION_FIELDS["unfriendly_reason"]),
        41: _value_from_row(row, SOURCE_QUESTION_FIELDS["receipt_address"]),
        42: _value_from_row(row, SOURCE_QUESTION_FIELDS["receipt_store_code"]),
        43: _value_from_row(row, SOURCE_QUESTION_FIELDS["think25"]),
        44: _value_from_row(row, SOURCE_QUESTION_FIELDS["underage_evidence"]),
        45: _value_from_row(row, SOURCE_QUESTION_FIELDS["underage_details"]),
        46: _value_from_row(row, SOURCE_QUESTION_FIELDS["unusual"]),
        47: _value_from_row(row, SOURCE_QUESTION_FIELDS["confirm_id"]),
    }
    return {col: (None if value is not None and isinstance(value, float) and pd.isna(value) else value) for col, value in values.items()}


def _row_values_to_sheet(ws, row_num, values, max_input_col=221):
    # Clear all source/input columns, including currently unused hidden question columns,
    # but do not touch the template's helper-formula columns at 222+.
    for col in range(1, max_input_col + 1):
        ws.cell(row_num, col).value = None
    for col, value in values.items():
        if value is not None and not (isinstance(value, float) and pd.isna(value)):
            ws.cell(row_num, col).value = value


def _existing_audit_ids(ws):
    ids = set()
    for r in range(4, ws.max_row + 1):
        audit_id = _normalise_text(ws.cell(r, 3).value)
        if audit_id:
            ids.add(audit_id.upper())
    return ids


def _last_data_row(ws):
    last = 3
    for r in range(4, ws.max_row + 1):
        if _normalise_text(ws.cell(r, 1).value):
            last = r
    return last


def _format_date_and_time(ws, row_num):
    for col in [5, 13, 14, 16]:
        ws.cell(row_num, col).number_format = "dd/mm/yyyy"
    ws.cell(row_num, 17).number_format = "hh:mm"


def _populate_store_list_from_master(wb, master_bytes: bytes):
    """Replace legacy external Core Stores lookups with values from the uploaded Master Schedule.

    The master schedule is the authoritative store source for this report. This prevents
    the client-facing export from inheriting #VALUE! errors caused by unavailable external links.
    """
    if "Store List - UPDATE" not in wb.sheetnames:
        return
    master_wb = load_workbook(io.BytesIO(master_bytes), data_only=True, read_only=True)
    try:
        master_ws = master_wb[master_wb.sheetnames[0]]
        stores = {}
        for row in master_ws.iter_rows(min_row=3, values_only=True):
            if len(row) < 11:
                continue
            epos = row[1]
            if epos is None:
                continue
            key = str(epos).strip().removesuffix(".0")
            address_parts = [str(v).strip() for v in row[6:11] if v is not None and str(v).strip()]
            stores[key] = {
                "region": row[0],
                "regional_manager": row[4],
                "area_manager": row[5],
                "people_partner": None,  # Not supplied as a name in the Master Schedule.
                "address": ", ".join(address_parts),
            }
        ws = wb["Store List - UPDATE"]
        for r in range(3, ws.max_row + 1):
            epos = ws.cell(r, 1).value
            if epos is None:
                continue
            key = str(epos).strip().removesuffix(".0")
            store = stores.get(key, {})
            # Columns B:F feed the existing report lookups. Replace the external formulas
            # with literal values from the supplied Master Schedule.
            ws.cell(r, 2).value = store.get("region")
            ws.cell(r, 3).value = store.get("regional_manager")
            ws.cell(r, 4).value = store.get("area_manager")
            ws.cell(r, 5).value = store.get("people_partner")
            ws.cell(r, 6).value = store.get("address")
    finally:
        master_wb.close()


def _set_full_recalc(wb):
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass


def build_report(template_bytes: bytes, csv_bytes: bytes, master_bytes: bytes, calendar_bytes: bytes,
                 financial_year: str, period: int):
    _, filtered, summary = filter_audits(csv_bytes, master_bytes, calendar_bytes, financial_year, period)
    schedule = summary.pop("schedule")
    trimmed = trim_template_xlsx(template_bytes)
    wb = load_workbook(io.BytesIO(trimmed), data_only=False, keep_links=True)
    if "This month" not in wb.sheetnames or "YTD" not in wb.sheetnames:
        raise ValueError("The report template must contain 'This month' and 'YTD' tabs.")
    this_ws = wb["This month"]
    ytd_ws = wb["YTD"]

    # This month is replaced in full.
    this_existing_last = _last_data_row(this_ws)
    this_start = 4
    this_count = len(filtered)
    this_end = max(this_existing_last, this_start + this_count - 1)
    prototype_this = 4 if this_ws.max_row >= 4 else 3
    this_max_row = this_ws.max_row
    this_max_col = this_ws.max_column
    for r in range(this_start, this_end + 1):
        if r > this_max_row:
            _copy_row_style_and_formulas(this_ws, prototype_this, r, this_max_col)
            this_max_row = r
        else:
            for c in range(1, 222):
                this_ws.cell(r, c).value = None
    for i, (_, row) in enumerate(filtered.iterrows()):
        r = this_start + i
        if r > this_max_row:
            _copy_row_style_and_formulas(this_ws, prototype_this, r, this_max_col)
            this_max_row = r
        sid = _normalise_text(row.get("site_internal_id")).upper()
        values = _make_output_row(row, schedule.get(sid), result_case="upper")
        _row_values_to_sheet(this_ws, r, values)
        _format_date_and_time(this_ws, r)
    this_ws["A2"] = f"=COUNTA(A4:A{max(45129, this_start + this_count - 1)})"

    # YTD retains its previous records and appends only audits not already present.
    existing_ids = _existing_audit_ids(ytd_ws)
    new_rows = []
    skipped_duplicate_ids = []
    for _, row in filtered.iterrows():
        audit_id = _normalise_text(row.get("internal_id")).upper()
        if audit_id and audit_id in existing_ids:
            skipped_duplicate_ids.append(audit_id)
            continue
        new_rows.append(row)
        if audit_id:
            existing_ids.add(audit_id)
    ytd_last = _last_data_row(ytd_ws)
    ytd_start_new = max(4, ytd_last + 1)
    prototype_ytd = ytd_last if ytd_last >= 4 else 4
    ytd_max_row = ytd_ws.max_row
    ytd_max_col = ytd_ws.max_column
    for i, row in enumerate(new_rows):
        r = ytd_start_new + i
        if r > ytd_max_row:
            _copy_row_style_and_formulas(ytd_ws, prototype_ytd, r, ytd_max_col)
            ytd_max_row = r
        else:
            for c in range(1, 222):
                ytd_ws.cell(r, c).value = None
        sid = _normalise_text(row.get("site_internal_id")).upper()
        values = _make_output_row(row, schedule.get(sid), result_case="upper")
        _row_values_to_sheet(ytd_ws, r, values)
        _format_date_and_time(ytd_ws, r)
    ytd_final_last = _last_data_row(ytd_ws)
    ytd_ws["A2"] = f"=COUNTA(A4:A{max(14721, ytd_final_last)})"
    summary["ytd_appended"] = len(new_rows)
    summary["ytd_duplicates_skipped"] = len(skipped_duplicate_ids)
    summary["duplicate_ids_skipped"] = skipped_duplicate_ids
    summary["this_month_rows"] = this_count
    summary["ytd_rows_after_append"] = max(0, ytd_final_last - 3)
    summary["period"] = int(period)
    summary["financial_year"] = financial_year

    # Populate the report's store lookup table from the supplied Master Schedule.
    # This removes legacy external Core Stores links before formula recalculation.
    _populate_store_list_from_master(wb, master_bytes)

    # Summary Data is a formula-driven row-for-row view of This month.
    # Keep only rows corresponding to actual current-period audits (headers end at row 7).
    summary_ws = wb["Summary Data"] if "Summary Data" in wb.sheetnames else None
    if summary_ws is not None:
        summary_data_end = 7 + this_count
        if summary_ws.max_row > summary_data_end:
            summary_ws.delete_rows(summary_data_end + 1, summary_ws.max_row - summary_data_end)

    _set_full_recalc(wb)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), summary


def make_client_workbook(live_workbook_bytes: bytes) -> bytes:
    """Create client-facing workbook containing Summary Data through Product Performance,
    with formula cells replaced by calculated values when LibreOffice is available.
    """
    import os
    import shutil
    import subprocess
    import tempfile

    # LibreOffice calculates the template's formulas and stores cached results.
    # This is important because openpyxl itself does not calculate Excel formulas.
    with tempfile.TemporaryDirectory(prefix="onestop_client_export_") as temp_dir:
        source_path = os.path.join(temp_dir, "live_input.xlsx")
        with open(source_path, "wb") as f:
            f.write(live_workbook_bytes)
        lo_dir = os.path.join(temp_dir, "recalculated")
        os.makedirs(lo_dir, exist_ok=True)
        soffice = shutil.which("libreoffice") or shutil.which("soffice")
        recalculated_path = None
        if soffice:
            profile = os.path.join(temp_dir, "lo_profile")
            try:
                subprocess.run(
                    [soffice, "-env:UserInstallation=file://" + profile, "--headless", "--convert-to", "xlsx", "--outdir", lo_dir, source_path],
                    check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180,
                )
                candidate = os.path.join(lo_dir, "live_input.xlsx")
                if os.path.exists(candidate):
                    recalculated_path = candidate
            except Exception:
                recalculated_path = None

        formula_wb = load_workbook(io.BytesIO(live_workbook_bytes), data_only=False, keep_links=False)
        value_wb = load_workbook(recalculated_path, data_only=True, keep_links=False) if recalculated_path else None
        try:
            start_idx = formula_wb.sheetnames.index("Summary Data")
            # Remove all internal/admin sheets before Summary Data.
            for name in formula_wb.sheetnames[:start_idx]:
                formula_wb.remove(formula_wb[name])
            # Only retain the client-facing run from Summary Data through Product Performance.
            wanted = ["Summary Data", "Store Performance", "Regional Performance",
                      "North Region Performance", "Central Region Performance",
                      "South Region Performance", "Product Performance"]
            for name in list(formula_wb.sheetnames):
                if name not in wanted:
                    formula_wb.remove(formula_wb[name])
            if value_wb:
                for ws in formula_wb.worksheets:
                    if ws.title not in value_wb.sheetnames:
                        continue
                    value_ws = value_wb[ws.title]
                    for row in ws.iter_rows():
                        for cell in row:
                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                cell.value = value_ws[cell.coordinate].value
            else:
                # A calculated-value client file cannot safely be produced without a formula engine.
                raise RuntimeError("Could not recalculate workbook formulas. Install LibreOffice to generate the client-facing values-only workbook.")
            result = io.BytesIO()
            formula_wb.save(result)
            return result.getvalue()
        finally:
            formula_wb.close()
            if value_wb:
                value_wb.close()


def _sheet_paths_from_workbook(zf):
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
    workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {}
    for rel in rels_root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = str(PurePosixPath("xl") / target)
        rel_map[rid] = path.replace("xl/xl/", "xl/")
    result = {}
    for sheet in workbook_root.findall(".//main:sheets/main:sheet", ns):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if name and rid in rel_map:
            result[name] = rel_map[rid]
    return result


def _active_data_row_and_trim(sheet_xml: bytes, keep_min=3):
    # Find the last row containing a non-empty column-A cell without parsing every
    # full row. This is much faster on the original template's million-row style tail.
    active_last = keep_min
    a_cell_pattern = re.compile(
        rb'<c\b(?=[^>]*\br="A(\d+)")[^>]*>(.*?)</c>', re.DOTALL
    )
    for match in a_cell_pattern.finditer(sheet_xml):
        row_match = re.search(rb'\br="A(\d+)"', match.group(0))
        if not row_match:
            continue
        row_num = int(row_match.group(1))
        if row_num >= 4 and (b"<v>" in match.group(2) or b"<is>" in match.group(2)):
            active_last = max(active_last, row_num)

    # Rows are stored in ascending order. Find the first row after the last active
    # record and cut the unused tail in one operation instead of processing a million
    # complete row elements.
    row_open_pattern = re.compile(rb'<row\b[^>]*\br="(\d+)"[^>]*>')
    cut_at = None
    for match in row_open_pattern.finditer(sheet_xml):
        if int(match.group(1)) > active_last:
            cut_at = match.start()
            break
    if cut_at is not None:
        sheet_data_end = sheet_xml.find(b"</sheetData>", cut_at)
        if sheet_data_end >= 0:
            cleaned = sheet_xml[:cut_at] + sheet_xml[sheet_data_end:]
        else:
            cleaned = sheet_xml[:cut_at]
    else:
        cleaned = sheet_xml

    # The supplied template has no data or headers in columns 48:221.
    # Remove only empty styled cells from that unused range to avoid loading hundreds
    # of thousands of blank cell objects. Any cell with a value/formula is retained.
    cell_pattern = re.compile(rb'<c\b[^>]*?/>|<c\b[^>]*>.*?</c>', re.DOTALL)
    def strip_empty_middle_cell(match):
        cell = match.group(0)
        ref_match = re.search(rb'\br="([A-Z]+)\d+"', cell)
        if not ref_match:
            return cell
        letters = ref_match.group(1).decode("ascii")
        col_num = 0
        for ch in letters:
            col_num = col_num * 26 + ord(ch) - 64
        if 48 <= col_num <= 221 and not any(tag in cell for tag in (b"<v>", b"<f", b"<is>")):
            return b""
        return cell
    cleaned = cell_pattern.sub(strip_empty_middle_cell, cleaned)

    # Update the sheet dimension to match the retained active range.
    dim = re.search(rb'<dimension\b[^>]*\bref="([^"]+)"', cleaned)
    if dim:
        old_ref = dim.group(1).decode("ascii", errors="ignore")
        end_col = "IP"
        m = re.search(r":([A-Z]+)\d+$", old_ref)
        if m:
            end_col = m.group(1)
        cleaned = cleaned[:dim.start(1)] + f"A1:{end_col}{active_last}".encode() + cleaned[dim.end(1):]
    return cleaned, active_last


def trim_template_xlsx(file_bytes: bytes) -> bytes:
    """Strip million-row style-only tails from the two data sheets before openpyxl loads them."""
    source = io.BytesIO(file_bytes)
    out = io.BytesIO()
    with zipfile.ZipFile(source, "r") as zin:
        paths = _sheet_paths_from_workbook(zin)
        targets = {"This month", "YTD"}
        replacements = {}
        for sheet_name in targets:
            path = paths.get(sheet_name)
            if not path or path not in zin.namelist():
                continue
            xml = zin.read(path)
            cleaned, _ = _active_data_row_and_trim(xml)
            replacements[path] = cleaned
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
            for item in zin.infolist():
                data = replacements.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
    return out.getvalue()
