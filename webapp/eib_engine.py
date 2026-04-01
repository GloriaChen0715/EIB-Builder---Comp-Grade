"""
Compensation Grade EIB Builder - Core Engine
Replicates the Excel workbook logic for generating Workday EIB templates.
"""
import math
import os
from datetime import date
from typing import List, Dict, Optional
import pandas as pd


# Location factors by employment type
FACTORS = {
    "Exempt or TTC": [round(0.90 + i * 0.01, 2) for i in range(31)],  # 0.90 to 1.20
    "Hourly": [round(0.90 + i * 0.01, 2) for i in range(31)],
    "Executive": [1.00],
    "Puerto Rico": [0.80],
    "TTC – Flat Incentive": [round(0.90 + i * 0.01, 2) for i in range(31)],  # 0.90 to 1.20, same as Exempt
}

CAREER_BANDS_EXECUTIVE = ["L01", "L02", "L03", "L04", "L05", "L06"]
CAREER_BANDS_FULL = ["L07", "L08", "L09", "L10", "L11", "SL07", "SL08", "SL09", "SL10", "SL11", "TMP"]


def smart_round(value: float, is_hourly: bool) -> float:
    """Round based on employment type and value magnitude.
    Hourly or values < 100: round to 2 decimal places.
    Otherwise: round to nearest 100.
    """
    if is_hourly or abs(value) < 100:
        return round(value, 2)
    else:
        return round(value / 100) * 100


def generate_eib_data(
    jobs: List[Dict],
    template_type: str,  # "New" or "Update"
    employment_type: str,  # "Exempt or TTC", "Hourly", "Executive", "Puerto Rico"
) -> pd.DataFrame:
    """
    Generate the full EIB output for all jobs.

    Each job dict has: job_code, job_title, career_band, national_market_50th, effective_date
    """
    factors = FACTORS.get(employment_type, FACTORS["Exempt or TTC"])
    is_hourly = employment_type == "Hourly"

    rows = []

    for job_idx, job in enumerate(jobs):
        job_code = str(job["job_code"]).strip()
        job_title = str(job["job_title"]).strip()
        career_band = str(job["career_band"]).strip()
        mid_national = float(job["national_market_50th"])
        eff_date = job.get("effective_date", date.today().replace(month=1, day=1).isoformat())

        # Base compensation grade ranges
        base_min = 0.85 * mid_national
        base_mid = mid_national
        base_max = 1.15 * mid_national

        # Round base values
        base_min_r = smart_round(base_min, is_hourly)
        base_mid_r = smart_round(base_mid, is_hourly)
        base_max_r = smart_round(base_max, is_hourly)

        grade_name = f"BAND {career_band} - {job_code} {job_title}"

        for factor_idx, factor in enumerate(factors):
            row = {}

            # --- Compensation Grade Data (only on first factor row) ---
            is_first = (factor_idx == 0)

            row["Spreadsheet Key*"] = job_code

            if is_first:
                if template_type == "Update":
                    row["Compensation Grade"] = f"Grade_{job_code}"
                else:
                    row["Compensation Grade"] = ""

                if template_type == "New":
                    row["Compensation Grade ID"] = f"GRADE_{job_code}"
                else:
                    row["Compensation Grade ID"] = ""

                row["Effective Date"] = eff_date
                row["Name*"] = grade_name
                row["Description"] = ""
                row["Compensation Element*+*"] = "PY_REGULAR_BASE_PAY"
                row["Eligibility Rule+"] = ""
                row["Number of Segments"] = 3
                row["Minimum"] = base_min_r
                row["Midpoint"] = base_mid_r
                row["Maximum"] = base_max_r
                row["Spread"] = ""
                row["Segment 1 Top"] = base_min_r
                row["Segment 2 Top"] = base_mid_r
                row["Segment 3 Top"] = base_max_r
                row["Segment 4 Top"] = ""
                row["Segment 5 Top"] = ""
                row["Currency"] = "USD"
                row["Frequency"] = "ANNUAL"
                row["Salary Plan"] = ""
                row["Allow Override"] = "N"
            else:
                for field in ["Compensation Grade", "Compensation Grade ID", "Effective Date",
                              "Name*", "Description", "Compensation Element*+*", "Eligibility Rule+",
                              "Number of Segments", "Minimum", "Midpoint", "Maximum", "Spread",
                              "Segment 1 Top", "Segment 2 Top", "Segment 3 Top", "Segment 4 Top",
                              "Segment 5 Top", "Currency", "Frequency", "Salary Plan", "Allow Override"]:
                    row[field] = ""

            # --- Compensation Grade Profile Data ---
            row["Row ID*"] = factor_idx + 1

            factor_str = f"{factor:.2f}"
            profile_id_prefix = f"GRADE_PROFILE_{factor_str}_{job_code}"

            if template_type == "Update":
                row["Compensation Grade Profile"] = profile_id_prefix
                row["Compensation Grade Profile ID"] = ""
            elif template_type == "New":
                row["Compensation Grade Profile"] = ""
                row["Compensation Grade Profile ID"] = profile_id_prefix

            row["Effective Date (Profile)"] = eff_date

            profile_name = f"{grade_name} PROFILE Location Factor - {factor_str}"
            row["Name* (Profile)"] = profile_name
            row["Description (Profile)"] = profile_name
            row["Compensation Element*+* (Profile)"] = "PY_REGULAR_BASE_PAY"

            # Eligibility rule for profile
            if employment_type == "Executive":
                row["Eligibility Rule+ (Profile)"] = "EMPLOYEES_-_DIRECTOR_&_ABOVE"
            else:
                row["Eligibility Rule+ (Profile)"] = f"LOCATION_FACTOR_-_{factor_str}"

            row["Inactive"] = ""

            # Profile pay ranges: factor * MID (national), then apply 0.85/1.15
            profile_mid = factor * mid_national
            profile_min = profile_mid * 0.85
            profile_max = profile_mid * 1.15

            profile_mid_r = smart_round(profile_mid, is_hourly)
            profile_min_r = smart_round(smart_round(profile_mid, is_hourly) * 0.85, is_hourly)
            profile_max_r = smart_round(smart_round(profile_mid, is_hourly) * 1.15, is_hourly)

            row["Number of Segments (Profile)"] = 3
            row["Minimum (Profile)"] = profile_min_r
            row["Midpoint (Profile)"] = profile_mid_r
            row["Maximum (Profile)"] = profile_max_r
            row["Spread (Profile)"] = ""
            row["Segment 1 Top (Profile)"] = profile_min_r
            row["Segment 2 Top (Profile)"] = profile_mid_r
            row["Segment 3 Top (Profile)"] = profile_max_r
            row["Segment 4 Top (Profile)"] = ""
            row["Segment 5 Top (Profile)"] = ""
            row["Currency (Profile)"] = "USD"
            row["Frequency (Profile)"] = "ANNUAL"
            row["Salary Plan (Profile)"] = ""
            row["Allow Override (Profile)"] = "N"

            rows.append(row)

    return pd.DataFrame(rows)


def generate_workday_eib(
    jobs: List[Dict],
    template_type: str,
    employment_type: str,
) -> pd.DataFrame:
    """
    Generate the final Workday EIB format matching the 'Compensation Grade' output sheet.
    Maps internal columns to the exact Workday EIB column structure.
    """
    df = generate_eib_data(jobs, template_type, employment_type)
    if df.empty:
        return df

    # Build the output in Workday EIB format
    output_rows = []
    factors = FACTORS.get(employment_type, FACTORS["Exempt or TTC"])
    is_hourly = employment_type == "Hourly"
    is_ttc_flat_incentive = employment_type == "TTC – Flat Incentive"

    for job_idx, job in enumerate(jobs):
        job_code = str(job["job_code"]).strip()
        job_title = str(job["job_title"]).strip()
        career_band = str(job["career_band"]).strip()
        mid_national = float(job["national_market_50th"])
        eff_date = job.get("effective_date", date.today().replace(month=1, day=1).isoformat())
        
        # CCI TTC: Get Customer Care Incentive amount
        cci_amount = 0
        if is_ttc_flat_incentive:
            cci_amount = float(job.get("customer_care_incentive", 0) or 0)

        # Base compensation grade ranges
        if is_ttc_flat_incentive:
            # For CCI TTC: Base TTC = National Market 50th + CCI (at factor 1.0)
            base_mid_r = smart_round(mid_national + cci_amount, False)
            base_min_r = smart_round(base_mid_r * 0.85, False)
            base_max_r = smart_round(base_mid_r * 1.15, False)
        else:
            base_min_r = smart_round(0.85 * mid_national, is_hourly)
            base_mid_r = smart_round(mid_national, is_hourly)
            base_max_r = smart_round(1.15 * mid_national, is_hourly)
        
        grade_name = f"BAND {career_band} - {job_code} {job_title}"

        for factor_idx, factor in enumerate(factors):
            row = {}
            is_first = (factor_idx == 0)
            factor_str = f"{factor:.2f}"

            # Column B: Spreadsheet Key
            row["Spreadsheet Key*"] = job_code
            row["Add Only"] = ""

            # Compensation Grade fields (only first row)
            if is_first:
                row["Compensation Grade"] = f"Grade_{job_code}" if template_type == "Update" else ""
                row["Compensation Grade ID"] = f"GRADE_{job_code}" if template_type == "New" else ""
                row["Effective Date"] = eff_date
                row["Name*"] = grade_name
                row["Description"] = ""
                row["Compensation Element*+*"] = "PY_REGULAR_BASE_PAY"
                row["Eligibility Rule+"] = ""
                row["Number of Segments"] = 3
                row["Minimum"] = base_min_r
                row["Midpoint"] = base_mid_r
                row["Maximum"] = base_max_r
                row["Spread"] = ""
                row["Segment 1 Top"] = base_min_r
                row["Segment 2 Top"] = base_mid_r
                row["Segment 3 Top"] = base_max_r
                row["Segment 4 Top"] = ""
                row["Segment 5 Top"] = ""
                row["Currency"] = "USD"
                row["Frequency"] = "ANNUAL"
                row["Salary Plan"] = ""
                row["Allow Override"] = "N"
            else:
                for f in ["Compensation Grade", "Compensation Grade ID", "Effective Date",
                          "Name*", "Description", "Compensation Element*+*", "Eligibility Rule+",
                          "Number of Segments", "Minimum", "Midpoint", "Maximum", "Spread",
                          "Segment 1 Top", "Segment 2 Top", "Segment 3 Top", "Segment 4 Top",
                          "Segment 5 Top", "Currency", "Frequency", "Salary Plan", "Allow Override"]:
                    row[f] = ""

            # Compensation Grade Profile
            row["Row ID*"] = factor_idx + 1
            row["Delete"] = ""
            row["Compensation Grade Profile"] = f"GRADE_PROFILE_{factor_str}_{job_code}" if template_type == "Update" else ""
            row["Compensation Grade Profile ID"] = f"GRADE_PROFILE_{factor_str}_{job_code}" if template_type == "New" else ""
            row["Effective Date (Profile)"] = eff_date

            profile_name = f"{grade_name} PROFILE Location Factor - {factor_str}"
            row["Name* (Profile)"] = profile_name
            row["Description (Profile)"] = profile_name
            row["Compensation Element*+* (Profile)"] = "PY_REGULAR_BASE_PAY"

            if employment_type == "Executive":
                row["Eligibility Rule+ (Profile)"] = "EMPLOYEES_-_DIRECTOR_&_ABOVE"
            else:
                row["Eligibility Rule+ (Profile)"] = f"LOCATION_FACTOR_-_{factor_str}"

            row["Inactive"] = ""

            # Profile ranges calculation
            if is_ttc_flat_incentive:
                # CCI TTC: Profile TTC Midpoint = (National Market 50th * factor) + CCI
                profile_mid_r = smart_round((factor * mid_national) + cci_amount, False)
                profile_min_r = smart_round(profile_mid_r * 0.85, False)
                profile_max_r = smart_round(profile_mid_r * 1.15, False)
            else:
                profile_mid_r = smart_round(factor * mid_national, is_hourly)
                profile_min_r = smart_round(profile_mid_r * 0.85, is_hourly)
                profile_max_r = smart_round(profile_mid_r * 1.15, is_hourly)

            row["Number of Segments (Profile)"] = 3
            row["Minimum (Profile)"] = profile_min_r
            row["Midpoint (Profile)"] = profile_mid_r
            row["Maximum (Profile)"] = profile_max_r
            row["Spread (Profile)"] = ""
            row["Segment 1 Top (Profile)"] = profile_min_r
            row["Segment 2 Top (Profile)"] = profile_mid_r
            row["Segment 3 Top (Profile)"] = profile_max_r
            row["Segment 4 Top (Profile)"] = ""
            row["Segment 5 Top (Profile)"] = ""
            row["Currency (Profile)"] = "USD"
            row["Frequency (Profile)"] = "ANNUAL"
            row["Salary Plan (Profile)"] = ""
            row["Allow Override (Profile)"] = "N"

            output_rows.append(row)

    return pd.DataFrame(output_rows)


def parse_uploaded_excel(file_path: str) -> tuple:
    """Parse an uploaded Excel file that matches the Reference Data input format.
    Returns (jobs_list, template_type, employment_type)
    """
    df = pd.read_excel(file_path, sheet_name=0, header=None)

    template_type = "New"
    employment_type = "Exempt or TTC"
    jobs = []

    # Try to find the config rows
    for idx, row in df.iterrows():
        val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if "EIB Template" in val:
            template_type = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "New"
        elif "Type of Employment" in val:
            employment_type = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "Exempt or TTC"

    # Find header row with "Job Code"
    header_idx = None
    for idx, row in df.iterrows():
        for col_idx, val in enumerate(row):
            if pd.notna(val) and "Job Code" in str(val):
                header_idx = idx
                break
        if header_idx is not None:
            break

    if header_idx is not None:
        data_df = pd.read_excel(file_path, sheet_name=0, header=header_idx)

        col_cci = _find_column(data_df.columns, _CCI_ALIASES)
        col_exp_inc = _find_column(data_df.columns, _EXPERIENCE_INCENTIVE_ALIASES)

        def _parse_cell_raw(val):
            if pd.isna(val):
                return None
            if isinstance(val, (int, float)):
                return float(val) if float(val) > 0 else None
            s = str(val).strip().replace("$", "").replace(",", "").replace(" ", "")
            if not s:
                return None
            try:
                n = float(s)
                return n if n > 0 else None
            except (ValueError, TypeError):
                return None

        for _, row in data_df.iterrows():
            job_code = row.get("Job Code.") or row.get("Job Code")
            if pd.isna(job_code) or str(job_code).strip() == "":
                continue

            flat_total = _sum_flat_incentives(row, col_cci, col_exp_inc, _parse_cell_raw)
            cci_value = flat_total if flat_total > 0 else None

            jobs.append({
                "job_code": str(int(job_code)) if isinstance(job_code, float) else str(job_code),
                "job_title": str(row.get("Job Title", "")) if pd.notna(row.get("Job Title", "")) else "",
                "career_band": str(row.get("Career Band", "")) if pd.notna(row.get("Career Band", "")) else "",
                "national_market_50th": float(row.get("National Market 50th", 0)) if pd.notna(row.get("National Market 50th")) else 0,
                "effective_date": str(row.get("Effective Date", date.today().replace(month=1, day=1).isoformat()))[:10],
                "customer_care_incentive": cci_value,
            })

    return jobs, template_type, employment_type


# Column name aliases for flexible matching in Job Code Table uploads
_JOB_CODE_ALIASES = ["job code", "job_code", "jobcode", "job code.", "job_code.", "job profile id", "job profile reference id"]
_JOB_TITLE_ALIASES = ["job title", "job_title", "jobtitle", "title", "job name", "job profile name", "job profile", "position title"]
_CAREER_BAND_ALIASES = ["career band", "career_band", "careerband", "band", "compensation band", "comp band", "management level", "job level", "level", "job family", "job family group", "grade"]
_MARKET_50TH_ALIASES = [
    "national market 50th", "national_market_50th", "market 50th", "market_50th",
    "midpoint", "mid", "national market midpoint", "50th percentile", "50th",
    "national 50th", "market midpoint",
]
_CCI_ALIASES = [
    "customer service incentive", "customer_service_incentive", "customer care incentive",
    "customer_care_incentive", "customer care incentives",
]
# Order: most specific first (partial-match fallback uses substring containment).
_EXPERIENCE_INCENTIVE_ALIASES = [
    "experience incentives target",
    "experience_incentives_target",
    "experience incentive target",
    "experience_incentive_target",
    "experience incentive",
    "experience_incentive",
    "experience incentives",
    "experience incentivs",  # typo variant in some exports
]


def _sum_flat_incentives(row, col_csi: str, col_exp: str, parse_cell) -> float:
    """Sum Customer Service and Experience incentive amounts (>= 0 each) for TTC flat midpoint."""
    total = 0.0
    for col in (col_csi, col_exp):
        if not col:
            continue
        v = parse_cell(row.get(col))
        if v is not None and v > 0:
            total += v
    return total


def _find_column(df_columns: List[str], aliases: List[str]) -> str:
    """Find a column name in a DataFrame by matching against a list of aliases (case-insensitive)."""
    lower_cols = {c.strip().lower(): c for c in df_columns}
    for alias in aliases:
        if alias in lower_cols:
            return lower_cols[alias]
    # Partial match fallback
    for alias in aliases:
        for lc, orig in lower_cols.items():
            if alias in lc:
                return orig
    return ""


def parse_job_code_table(file_path: str) -> dict:
    """Parse a Job Code Table Excel/CSV file.
    Returns a dict with:
      - jobs: list of {job_code, job_title, career_band, current_market_50th}
      - columns_found: list of all column names in the file
      - mapped: dict of which columns were auto-matched
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(file_path)
    else:
        # Try to read; if first row isn't header, scan for header row
        df = pd.read_excel(file_path, sheet_name=0, header=None)
        header_idx = None
        for idx, row in df.iterrows():
            row_str = " ".join(str(v).lower() for v in row if pd.notna(v))
            # Look for rows that have multiple column-like headers, not just title rows
            cell_count = sum(1 for v in row if pd.notna(v) and len(str(v).strip()) > 0)
            # A header row should have multiple cells and contain "job code" or similar
            if cell_count >= 3 and ("job code" in row_str or "job_code" in row_str or "job profile" in row_str):
                header_idx = idx
                break
        if header_idx is not None:
            df = pd.read_excel(file_path, sheet_name=0, header=header_idx)
        else:
            df = pd.read_excel(file_path, sheet_name=0, header=0)

    all_columns = [str(c) for c in df.columns]

    # Find columns
    col_job_code = _find_column(df.columns, _JOB_CODE_ALIASES)
    col_job_title = _find_column(df.columns, _JOB_TITLE_ALIASES)
    col_career_band = _find_column(df.columns, _CAREER_BAND_ALIASES)
    col_market_50th = _find_column(df.columns, _MARKET_50TH_ALIASES)
    col_cci = _find_column(df.columns, _CCI_ALIASES)
    col_exp_inc = _find_column(df.columns, _EXPERIENCE_INCENTIVE_ALIASES)

    mapped = {
        "job_code": col_job_code or None,
        "job_title": col_job_title or None,
        "career_band": col_career_band or None,
        "current_market_50th": col_market_50th or None,
        "customer_service_incentive": col_cci or None,
        "experience_incentive": col_exp_inc or None,
    }

    if not col_job_code:
        raise ValueError(
            f"Could not find a 'Job Code' column in the uploaded file. "
            f"Columns found: {', '.join(all_columns)}"
        )

    def parse_numeric(val):
        """Parse a value that might be numeric, handling various formats."""
        if pd.isna(val):
            return None
        if isinstance(val, (int, float)):
            return float(val) if val > 0 else None
        # Handle string formats: "$1,234.56", "1,234", etc.
        val_str = str(val).strip()
        if not val_str:
            return None
        # Remove currency symbols, commas, spaces
        val_str = val_str.replace('$', '').replace(',', '').replace(' ', '')
        try:
            num = float(val_str)
            return num if num > 0 else None
        except (ValueError, TypeError):
            return None

    # Skip any completely blank rows at the start
    df = df.dropna(how='all')

    jobs = []

    for _, row in df.iterrows():
        job_code = row.get(col_job_code) if col_job_code else None
        if pd.isna(job_code) or str(job_code).strip() == "":
            continue

        # Convert job code to clean string
        if isinstance(job_code, float):
            job_code = str(int(job_code))
        else:
            job_code = str(job_code).strip()

        job_title = ""
        if col_job_title and pd.notna(row.get(col_job_title)):
            job_title = str(row[col_job_title]).strip()

        career_band = ""
        if col_career_band and pd.notna(row.get(col_career_band)):
            career_band = str(row[col_career_band]).strip()

        current_market_50th = parse_numeric(row.get(col_market_50th)) if col_market_50th else None

        # Sum Customer Service + Experience incentives (flat $ for TTC midpoint)
        flat_total = _sum_flat_incentives(row, col_cci, col_exp_inc, parse_numeric)
        customer_care_incentive = flat_total if flat_total > 0 else None

        jobs.append({
            "job_code": job_code,
            "job_title": job_title,
            "career_band": career_band,
            "current_market_50th": current_market_50th,
            "customer_care_incentive": customer_care_incentive,
        })

    return {
        "jobs": jobs,
        "columns_found": all_columns,
        "mapped": mapped,
    }
