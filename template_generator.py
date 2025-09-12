from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime, date
from typing import Dict, List, Union

NumberLike = Union[float, int, str]

def _clean_to_float(v: NumberLike) -> float:
    """Convert numbers or currency-like strings to float; invalid -> 0.0"""
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace("$", "").replace(",", "").strip())
        except Exception:
            return 0.0
    return 0.0

def _coerce_year_key(k) -> Union[int, None]:
    """Coerce year keys like '2025', '2025.0', 2025.0 -> 2025; return None if impossible."""
    if isinstance(k, int):
        return k
    if isinstance(k, float):
        return int(round(k))
    if isinstance(k, str):
        s = k.strip()
        try:
            return int(s)
        except Exception:
            try:
                return int(float(s))
            except Exception:
                return None
    return None

def _get_monthly_premium_frontfill_tail(
    mp_dict: Dict[int, List[NumberLike]],
    year: int,
    month_1based: int
) -> float:
    """
    Return the premium for (year, month). Each year maps to a list of monthly values.
    If a year has < 12 entries, assume the missing months are at the FRONT (Jan..),
    i.e., the provided values align to the END of the year (… Sep, Oct, Nov, Dec).
    Example: 9 entries -> treat Jan/Feb/Mar as 0, entries correspond to Apr..Dec.
    """
    lst = mp_dict.get(year)
    if not isinstance(lst, (list, tuple)):
        return 0.0

    clean = [_clean_to_float(x) for x in lst]
    n = len(clean)

    # Clamp month to [1..12]
    m = max(1, min(12, month_1based))

    if n >= 12:
        return clean[m - 1]

    # Front-fill zeros: provided values occupy the LAST n months of the year
    offset = 12 - n  # number of missing months at the start (treated as zeros)
    idx = (m - 1) - offset
    if idx < 0 or idx >= n:
        return 0.0
    return clean[idx]

def generate_return_template(
    insured_name: str,
    dob: str,
    carrier: str,
    le_months: int,
    le_report_date: str,
    death_benefit: float,
    investment: float,
    monthly_premiums: Dict[int, List[float]],
    output_filename: str
) -> str:
    # Parse dates / anchors
    dob_dt = datetime.strptime(dob, "%Y-%m-%d")
    le_report_dt = datetime.strptime(le_report_date, "%Y-%m-%d")
    today = date.today()

    # Elapsed/remaining LE
    elapsed_months = (today.year - le_report_dt.year) * 12 + today.month - le_report_dt.month
    remaining_le_months = max(le_months - elapsed_months, 0)
    remaining_le_years = (remaining_le_months + 11) // 12
    total_years = remaining_le_years + 3
    start_year = today.year

    # Approximate age at "today" anchored from LE report date + elapsed months
    age = int((le_report_dt - dob_dt).days / 365.25 + elapsed_months / 12)

    # Coerce year keys robustly and keep only valid ones
    monthly_premiums = {
        yk: v for k, v in (monthly_premiums or {}).items()
        if (yk := _coerce_year_key(k)) is not None
    }

    # Annual premium totals (robust to string inputs)
    annual_premiums: Dict[int, float] = {}
    for year, months in monthly_premiums.items():
        if not isinstance(months, (list, tuple)):
            annual_premiums[year] = 0.0
            continue
        annual_premiums[year] = sum(_clean_to_float(x) for x in months)

    # Load template and clear old output rows (keep headers up to row 6)
    wb = load_workbook("return_template_output.xlsx")
    ws = wb.active

    for _ in range(7, ws.max_row + 1):
        ws.delete_rows(7)

    # Header cells
    ws["B1"] = insured_name
    ws["B2"] = f"AGE: {age}"
    ws["B3"] = f"CARRIER: {carrier}"
    ws["E2"] = f"{remaining_le_months} MONTHS"
    ws["E3"] = death_benefit
    ws["E4"] = investment

    # === Auto-calc: next 3 months of premiums (including this month) -> E5 ===
    now_dt = datetime.now()
    cur_y, cur_m = now_dt.year, now_dt.month  # 1..12

    # Build (year, month) positions for current month + next two months, handling year wrap
    positions = []
    for off in range(3):
        y = cur_y + ((cur_m - 1 + off) // 12)
        m = ((cur_m - 1 + off) % 12) + 1
        positions.append((y, m))

    next_three_sum = sum(
        _get_monthly_premium_frontfill_tail(monthly_premiums, y, m) for (y, m) in positions
    )
    ws["E5"] = next_three_sum
    ws["E5"].number_format = '"$"#,##0.00'
    # === end auto-calc E5 ===

    # Year-by-year table
    cumulative = 0.0
    for i in range(total_years):
        year = start_year + i
        premium = float(annual_premiums.get(year, 0.0))
        cumulative += premium
        total_cost = float(investment) + cumulative
        profit = float(death_benefit) - total_cost
        simple_return = (profit / total_cost) if total_cost else 0.0
        acc_return = (simple_return / (i + 1)) if (i + 1) else 0.0

        marker = ""
        if i == remaining_le_years - 1:
            marker = "LE"
        elif i == remaining_le_years:
            marker = "LE+1"
        elif i == remaining_le_years + 1:
            marker = "LE+2"
        elif i == remaining_le_years + 2:
            marker = "LE+3"

        row = [
            year,
            premium,
            cumulative,
            total_cost,
            profit,
            simple_return,
            acc_return,
            marker
        ]
        ws.append(row)

        # Highlight LE row (bold + light blue fill)
        if i == remaining_le_years - 1:
            for col in range(2, 9):  # columns B..H
                cell = ws.cell(row=6 + i + 1, column=col)  # == row 7 + i
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

    # Number formats for the appended rows
    for row in range(7, 7 + total_years):
        for col in range(2, 6):  # B..E currency
            ws.cell(row=row, column=col).number_format = '"$"#,##0.00'
        for col in range(6, 8):  # F..G percentages
            ws.cell(row=row, column=col).number_format = '0.00%'

    # Copy the style from B6 into column A for all output rows
    ref_style = ws["B6"]._style
    for row in range(7, 7 + total_years):
        ws.cell(row=row, column=1)._style = ref_style

    # Clear any "LE Marker" header label in H6 to keep header area clean
    if ws["H6"].value == "LE Marker":
        ws["H6"].value = ""

    wb.save(output_filename)
    return output_filename
