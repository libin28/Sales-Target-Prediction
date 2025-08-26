
import io
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import plotly.express as px

st.set_page_config(page_title="Area-wise Sales Forecast", layout="wide")

# Add caching for better performance
@st.cache_data
def read_excel_cached(file_bytes, sheet_name=None, header=None, engine="openpyxl"):
    """Cached Excel reading function"""
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header, engine=engine)

@st.cache_data
def process_data_cached(file_bytes, sheet_name, header_row):
    """Cached data processing function"""
    try:
        if sheet_name.endswith('.xls'):
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row, engine="openpyxl")
        return df
    except Exception as e:
        st.error(f"Error reading sheet {sheet_name}: {e}")
        return None

# ----------------------
# Helpers
# ----------------------
MONTHS_MAP = {
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept":9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
}

def normalize_cols(df):
    # Lowercase, strip, remove extra spaces
    df.columns = [c.strip() for c in df.columns]
    df.columns = [c.replace("\n"," ").replace("\r"," ") for c in df.columns]
    df.columns = [' '.join(c.split()) for c in df.columns]
    return df

def convert_to_lakhs(value):
    """Convert value from rupees to lakhs (1 lakh = 100,000 rupees)"""
    if pd.isna(value) or value == 0:
        return value
    return value / 100000

def format_lakhs(value):
    """Format value in lakhs with proper notation"""
    if pd.isna(value) or value == 0:
        return "₹0"
    return f"₹{value:,.1f}L"

def process_dashboard_data(uploaded):
    """Process data for the sales comparison dashboard"""
    try:
        # Read the comparison report sheet
        comparison_df = pd.read_excel(uploaded, sheet_name='COMPARISON REPORT', header=6)

        # Find different sales sections and extract totals directly
        route_sales_total = {}
        msd_sales_total = {}
        inter_unit_total = {}

        current_section = None

        for i, row in comparison_df.iterrows():
            particulars = str(row.iloc[0]).strip().upper()

            if 'ROUTE SALES' in particulars:
                current_section = 'route'
                # Look for the total row for route sales
                for j in range(i+1, min(i+20, len(comparison_df))):
                    if j < len(comparison_df):
                        next_particulars = str(comparison_df.iloc[j, 0]).strip().upper()
                        if 'TOTAL' in next_particulars and 'ROUTE' in next_particulars:
                            # Extract year totals
                            year_cols = [col for col in comparison_df.columns if '-' in str(col) and any(char.isdigit() for char in str(col))]
                            for year_col in year_cols:
                                value = comparison_df.iloc[j][year_col]
                                if pd.notna(value):
                                    route_sales_total[year_col] = pd.to_numeric(value, errors='coerce')
                            break

            elif 'MSD SALES' in particulars:
                current_section = 'msd'
                # Look for the total row for MSD sales
                for j in range(i+1, min(i+20, len(comparison_df))):
                    if j < len(comparison_df):
                        next_particulars = str(comparison_df.iloc[j, 0]).strip().upper()
                        if 'TOTAL' in next_particulars and 'MSD' in next_particulars:
                            # Extract year totals
                            year_cols = [col for col in comparison_df.columns if '-' in str(col) and any(char.isdigit() for char in str(col))]
                            for year_col in year_cols:
                                value = comparison_df.iloc[j][year_col]
                                if pd.notna(value):
                                    msd_sales_total[year_col] = pd.to_numeric(value, errors='coerce')
                            break

            elif 'INTER UNIT' in particulars or 'INTER-UNIT' in particulars:
                current_section = 'inter_unit'
                # Look for the total row for Inter Unit sales
                for j in range(i+1, min(i+20, len(comparison_df))):
                    if j < len(comparison_df):
                        next_particulars = str(comparison_df.iloc[j, 0]).strip().upper()
                        if 'TOTAL' in next_particulars and ('INTER' in next_particulars or 'UNIT' in next_particulars):
                            # Extract year totals
                            year_cols = [col for col in comparison_df.columns if '-' in str(col) and any(char.isdigit() for char in str(col))]
                            for year_col in year_cols:
                                value = comparison_df.iloc[j][year_col]
                                if pd.notna(value):
                                    inter_unit_total[year_col] = pd.to_numeric(value, errors='coerce')
                            break

        return {
            'route_sales_total': route_sales_total,
            'msd_sales_total': msd_sales_total,
            'inter_unit_total': inter_unit_total
        }

    except Exception as e:
        st.error(f"Error processing dashboard data: {e}")
        return None

def create_dashboard_table(currency_format):
    """Create the dashboard table with the exact data provided"""
    try:
        # Exact data as provided by the user
        data = {
            'Column1': ['ROUTE SALES', 'MSD SALES', 'INTER UNIT SALES', 'Total'],
            '2018-2019': [132011864, 60767454, 28080085, 220859403],
            '2019-2020': [147473198, 61030939, 26841135, 235345272],
            '2020-2021': [195564515, 35538503, 22048038, 253151056],
            '2021-2022': [174604844, 30756095, 21108102, 226469041],
            '2022-2023': [167861540, 46372021, 28564074, 242797635],
            '2023-2024': [155908390, 41520083, 29965624, 227394097],
            '2024-2025': [144241963, 42410753, 28855386, 215508102],
            '2025-2026': [48043782, 14184331, 14569952, 76798065]
        }

        # Create DataFrame
        dashboard_df = pd.DataFrame(data)

        # Format the values based on currency selection
        year_columns = [col for col in dashboard_df.columns if col != 'Column1']

        for col in year_columns:
            if currency_format == 'Lakhs (₹L)':
                # Convert to lakhs and format with commas
                dashboard_df[col] = dashboard_df[col].apply(lambda x: f"{x/100000:,.0f}")
            else:
                # Convert to millions and format with commas
                dashboard_df[col] = dashboard_df[col].apply(lambda x: f"{x/10000000:,.2f}")

        # Rename the first column
        dashboard_df = dashboard_df.rename(columns={'Column1': 'Sales Type'})

        return dashboard_df

    except Exception as e:
        st.error(f"Error creating dashboard table: {e}")
        return None

def prepare_chart_data(sales_type, currency_format):
    """Prepare data for charts based on sales type selection using exact data"""
    try:
        # Exact data as provided
        years = ['2018-2019', '2019-2020', '2020-2021', '2021-2022', '2022-2023', '2023-2024', '2024-2025', '2025-2026']
        route_sales = [132011864, 147473198, 195564515, 174604844, 167861540, 155908390, 144241963, 48043782]
        msd_sales = [60767454, 61030939, 35538503, 30756095, 46372021, 41520083, 42410753, 14184331]
        inter_unit_sales = [28080085, 26841135, 22048038, 21108102, 28564074, 29965624, 28855386, 14569952]

        chart_data = []

        if sales_type == 'All':
            # Create data for all sales types
            for i, year in enumerate(years):
                if currency_format == 'Lakhs (₹L)':
                    chart_data.append({'Year': year, 'Sales_Category': 'Route Sales', 'Value': route_sales[i]/100000})
                    chart_data.append({'Year': year, 'Sales_Category': 'MSD Sales', 'Value': msd_sales[i]/100000})
                    chart_data.append({'Year': year, 'Sales_Category': 'Inter Unit Sales', 'Value': inter_unit_sales[i]/100000})
                else:
                    chart_data.append({'Year': year, 'Sales_Category': 'Route Sales', 'Value': route_sales[i]/10000000})
                    chart_data.append({'Year': year, 'Sales_Category': 'MSD Sales', 'Value': msd_sales[i]/10000000})
                    chart_data.append({'Year': year, 'Sales_Category': 'Inter Unit Sales', 'Value': inter_unit_sales[i]/10000000})

        elif sales_type == 'Route Sales':
            for i, year in enumerate(years):
                display_value = route_sales[i]/100000 if currency_format == 'Lakhs (₹L)' else route_sales[i]/10000000
                chart_data.append({'Year': year, 'Value': display_value})

        elif sales_type == 'MSD Sales':
            for i, year in enumerate(years):
                display_value = msd_sales[i]/100000 if currency_format == 'Lakhs (₹L)' else msd_sales[i]/10000000
                chart_data.append({'Year': year, 'Value': display_value})

        else:  # Inter Unit Sales
            for i, year in enumerate(years):
                display_value = inter_unit_sales[i]/100000 if currency_format == 'Lakhs (₹L)' else inter_unit_sales[i]/10000000
                chart_data.append({'Year': year, 'Value': display_value})

        return pd.DataFrame(chart_data)

    except Exception as e:
        st.error(f"Error preparing chart data: {e}")
        return pd.DataFrame()

def guess_month_cols(df):
    month_cols = []
    for c in df.columns:
        key = str(c).strip().lower()
        key = key.replace('.', ' ').replace('-', ' ').replace('_', ' ')
        key = ' '.join(key.split())

        # Check for exact month matches
        for m in MONTHS_MAP:
            if key == m or key.endswith(' '+m) or key.startswith(m+' '):
                month_cols.append(c)
                break
        else:
            # Check for partial matches (for columns like "April", "May", etc.)
            for m in MONTHS_MAP:
                if m in key and len(key) <= 15:  # Avoid matching long text
                    month_cols.append(c)
                    break

    # De-duplicate while preserving order
    seen = set()
    ordered = []
    for c in month_cols:
        lc = str(c).lower()
        if lc not in seen:
            ordered.append(c)
            seen.add(lc)

    # Sort by fiscal order Apr..Mar
    def month_order(cname):
        k = str(cname).strip().lower()
        parts = k.split()
        # try last token first
        tokens = [parts[-1]] + parts[:-1]
        for t in tokens:
            if t in MONTHS_MAP:
                return MONTHS_MAP[t]
        # fallback - try partial match
        for m in MONTHS_MAP:
            if m in k:
                return MONTHS_MAP[m]
        return 13

    ordered.sort(key=month_order)
    # Reorder to Apr..Mar cycle
    ordered = sorted(ordered, key=lambda c: (month_order(c)+8)%12)
    return ordered

def detect_keys(df):
    # Try to find Area, State, Year-like columns
    candidates = {c.lower(): c for c in df.columns}
    area = None
    state = None
    year = None
    for c in df.columns:
        cl = c.lower()
        if area is None and any(k in cl for k in ["area","region","territory","zone"]):
            area = c
        if state is None and "state" in cl:
            state = c
        if year is None and any(k in cl for k in ["year", "fy", "financial year", "fiscal", "yr"]):
            year = c
    return area, state, year

def to_fiscal_year_start(ycol_value):
    # Accept formats like "2018-19", "2018-2019", 2018, "FY 2018-19" etc.
    s = str(ycol_value).strip()
    # Try extract 4-digit year first occurrence
    import re
    m = re.search(r'(20\d{2})', s)
    if m:
        y = int(m.group(1))
    else:
        # fallback: if number-like convert
        try:
            y = int(float(s))
            if y < 2000: y = 2000 + (y % 100)  # crude fallback
        except:
            y = None
    if y is None:
        return None
    # Fiscal year assumed Apr..Mar, so start is April of detected y
    return pd.Timestamp(year=y, month=4, day=1)

def process_monthly_comparison_sheet(uploaded, sheet_name):
    """Process Monthly Comparison sheets with area-wise data blocks."""
    try:
        # Read the sheet with header at row 8 (where months are)
        if uploaded.name.endswith('.xls'):
            df = pd.read_excel(uploaded, sheet_name=sheet_name, header=8)
        else:
            df = pd.read_excel(uploaded, sheet_name=sheet_name, header=8, engine="openpyxl")

        # Define area patterns to look for
        area_patterns = ['KERALA', 'KARNATAKA', 'TAMIL NADU', 'OTHER STATES', 'MSD INSIDE KERALA', 'MSD OUTSIDE KERALA']

        all_data = []
        current_area = "KERALA"  # First section is typically Kerala

        # Month columns (excluding YEAR and TOTAL)
        month_cols = ['APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER',
                     'OCTOBER', 'NOVEMBER', 'DECEMBER', 'JANUARY', 'FEBRUARY', 'MARCH']

        for i, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().upper()

            # Check if this row indicates a new area
            for area in area_patterns:
                if area in first_val:
                    current_area = area
                    break

            # Check if this row contains year data (like 2018-2019, 2019-2020, etc.)
            if '-' in first_val and any(char.isdigit() for char in first_val):
                year = first_val

                # Extract monthly data for this year and area
                for month_col in month_cols:
                    if month_col in df.columns:
                        value = row[month_col]
                        if pd.notna(value) and value != 0:
                            # Convert month name to number
                            month_num = MONTHS_MAP.get(month_col.lower(), None)
                            if month_num:
                                # Create fiscal year date
                                fy_start = to_fiscal_year_start(year)
                                if fy_start:
                                    if month_num >= 4:
                                        dt = pd.Timestamp(year=fy_start.year, month=month_num, day=1)
                                    else:
                                        dt = pd.Timestamp(year=fy_start.year + 1, month=month_num, day=1)
                                else:
                                    dt = pd.NaT

                                all_data.append({
                                    "Area": current_area,
                                    "State": current_area if current_area in ['KERALA', 'KARNATAKA', 'TAMIL NADU'] else None,
                                    "FY": year,
                                    "MonthName": month_col,
                                    "Date": dt,
                                    "Month": month_num,
                                    "Sales": pd.to_numeric(value, errors="coerce"),
                                    "SourceSheet": sheet_name
                                })

        if all_data:
            result_df = pd.DataFrame(all_data)
            result_df = result_df.dropna(subset=["Sales"])
            return [result_df]
        else:
            return []

    except Exception as e:
        st.error(f"Error processing {sheet_name}: {e}")
        return []

def process_territory_data_from_yearly_sheets(uploaded, sheet_name):
    """Process yearly sheets to extract territory-level data."""

    # Define the Kerala territories (exact names as requested)
    kerala_territories = [
        'TRIVANDRUM', 'NEYYATINKARA', 'KOLLAM', 'PATHANAMTHITTA', 'KOTTAYAM',
        'ALAPPUZHA', 'IDUKKI', 'MOOVATTUPUZHA', 'ERNAMKULAM', 'PALAKKAD',
        'THRISSUR', 'EDAPAL', 'MALAPPURAM', 'KOZHIKODE CITY', 'VADAKARA',
        'WAYANAD', 'THALASSERY', 'KANNUR', 'KASARGOD'
    ]

    try:
        # Read the sheet with header at row 6 (where months are)
        if uploaded.name.endswith('.xls'):
            df = pd.read_excel(uploaded, sheet_name=sheet_name, header=6)
        else:
            df = pd.read_excel(uploaded, sheet_name=sheet_name, header=6, engine="openpyxl")

        all_data = []

        # Month columns (use the actual column names from the Excel)
        month_cols = ['April', 'May', 'June', 'July', 'August', 'September',
                     'October', 'November', 'December', 'January', 'February', 'March']

        # Extract fiscal year from sheet name
        fiscal_year = sheet_name

        # Find Route Sales section first
        route_sales_start = None
        for i, row in df.iterrows():
            particulars = str(row.iloc[0]).strip().upper()
            if 'ROUTE SALES' in particulars:
                route_sales_start = i
                break

        if route_sales_start is None:
            return []

        # Process Route Sales section to find territory data
        for i in range(route_sales_start + 1, min(route_sales_start + 100, len(df))):
            if i >= len(df):
                break

            area_name = str(df.iloc[i, 0]).strip().upper()

            # Skip empty rows and section headers
            if pd.isna(df.iloc[i, 0]) or area_name in ['NAN', '', 'TOTAL', 'INSIDE KERALA', 'CENTRAL ZONE', 'NORTH ZONE', 'SOUTH ZONE']:
                continue

            # Extract Kerala territory name
            territory_name = None
            for kerala_territory in kerala_territories:
                # Check for exact matches or aliases
                if kerala_territory == 'TRIVANDRUM' and any(alias in area_name for alias in ['TRIVANDRUM', 'TVM']):
                    territory_name = kerala_territory
                    break
                elif kerala_territory == 'ERNAMKULAM' and 'ERNAKULAM' in area_name:
                    territory_name = kerala_territory
                    break
                elif kerala_territory == 'KASARGOD' and 'KASARGODE' in area_name:
                    territory_name = kerala_territory
                    break
                elif kerala_territory == 'EDAPAL' and 'EDAPPAL' in area_name:
                    territory_name = kerala_territory
                    break
                elif kerala_territory in area_name:
                    territory_name = kerala_territory
                    break

            if territory_name:
                # Extract monthly data for this territory
                for month_col in month_cols:
                    if month_col in df.columns:
                        value = df.iloc[i][month_col]
                        if pd.notna(value) and value != 0:
                            try:
                                # Convert month name to number
                                month_num = MONTHS_MAP.get(month_col.lower(), None)
                                if month_num:
                                    # Create fiscal year date
                                    fy_start = to_fiscal_year_start(fiscal_year)
                                    if fy_start:
                                        if month_num >= 4:
                                            dt = pd.Timestamp(year=fy_start.year, month=month_num, day=1)
                                        else:
                                            dt = pd.Timestamp(year=fy_start.year + 1, month=month_num, day=1)
                                    else:
                                        dt = pd.NaT

                                    all_data.append({
                                        "Area": territory_name,
                                        "State": "KERALA",
                                        "FY": fiscal_year,
                                        "MonthName": month_col,
                                        "Date": dt,
                                        "Month": month_num,
                                        "Sales": pd.to_numeric(value, errors="coerce"),
                                        "SourceSheet": sheet_name
                                    })
                            except:
                                continue

        # Also search in DEBTORS section for missing territories (like NEYYATTINKARA)
        missing_territories = set(kerala_territories) - set([d['Area'] for d in all_data])
        if missing_territories:

            # Search entire sheet for missing territories
            for i, row in df.iterrows():
                area_name = str(row.iloc[0]).strip().upper()

                # Skip empty rows
                if pd.isna(row.iloc[0]) or area_name in ['NAN', 'PARTICULARS']:
                    continue

                # Check for missing territories in DEBTORS format
                territory_name = None
                for missing_territory in missing_territories:
                    if missing_territory == 'NEYYATINKARA' and 'NEYYATTINKARA' in area_name:
                        territory_name = 'NEYYATINKARA'
                        break
                    elif missing_territory == 'EDAPAL' and 'EDAPPAL' in area_name:
                        territory_name = 'EDAPAL'
                        break
                    elif missing_territory in area_name:
                        territory_name = missing_territory
                        break

                if territory_name:
                    # Extract monthly data for this territory
                    for month_col in month_cols:
                        if month_col in df.columns:
                            value = row[month_col]
                            if pd.notna(value) and value != 0:
                                try:
                                    # Convert month name to number
                                    month_num = MONTHS_MAP.get(month_col.lower(), None)
                                    if month_num:
                                        # Create fiscal year date
                                        fy_start = to_fiscal_year_start(fiscal_year)
                                        if fy_start:
                                            if month_num >= 4:
                                                dt = pd.Timestamp(year=fy_start.year, month=month_num, day=1)
                                            else:
                                                dt = pd.Timestamp(year=fy_start.year + 1, month=month_num, day=1)
                                        else:
                                            dt = pd.NaT

                                        all_data.append({
                                            "Area": territory_name,
                                            "State": "KERALA",
                                            "FY": fiscal_year,
                                            "MonthName": month_col,
                                            "Date": dt,
                                            "Month": month_num,
                                            "Sales": pd.to_numeric(value, errors="coerce"),
                                            "SourceSheet": sheet_name
                                        })
                                except:
                                    continue

        if all_data:
            result_df = pd.DataFrame(all_data)
            result_df = result_df.dropna(subset=["Sales"])
            return [result_df]
        else:
            return []

    except Exception as e:
        st.error(f"Error processing {sheet_name}: {e}")
        return []

def process_comparison_report_sheet(uploaded, sheet_name):
    """Process Comparison Report sheet with yearly data across areas."""
    try:
        # Read the sheet with header at row 6 (where years are)
        if uploaded.name.endswith('.xls'):
            df = pd.read_excel(uploaded, sheet_name=sheet_name, header=6)
        else:
            df = pd.read_excel(uploaded, sheet_name=sheet_name, header=6, engine="openpyxl")

        all_data = []

        # The years are in columns 1-8 (2018-2019, 2019-2020, etc.)
        year_columns = []
        for col in df.columns:
            col_str = str(col).strip()
            if '-' in col_str and any(char.isdigit() for char in col_str):
                year_columns.append(col)



        # Get the current month from the sheet (it shows "MONTH: JULY" etc.)
        current_month = "JULY"  # Default

        # Try to extract the actual month from the sheet
        for col in df.columns:
            col_str = str(col).upper().strip()
            for month_name, month_num_temp in MONTHS_MAP.items():
                if month_name.upper() in col_str:
                    current_month = month_name.upper()
                    break

        month_num = MONTHS_MAP.get(current_month.lower(), 7)


        # Process each row that contains area data
        for i, row in df.iterrows():
            area_name = str(row.iloc[0]).strip()

            # Skip empty rows or header rows
            if pd.isna(row.iloc[0]) or area_name in ['NaN', 'ROUTE SALES', 'Particulars']:
                continue

            # Clean up area name
            if area_name.startswith('Debtors - '):
                area_name = area_name.replace('Debtors - ', '').strip()

            # Extract data for each year
            for year_col in year_columns:
                value = row[year_col]
                if pd.notna(value) and value != 0:
                    try:
                        # Create fiscal year date
                        fy_start = to_fiscal_year_start(year_col)
                        if fy_start:
                            if month_num >= 4:
                                dt = pd.Timestamp(year=fy_start.year, month=month_num, day=1)
                            else:
                                dt = pd.Timestamp(year=fy_start.year + 1, month=month_num, day=1)
                        else:
                            dt = pd.NaT

                        all_data.append({
                            "Area": area_name,
                            "State": None,
                            "FY": year_col,
                            "MonthName": current_month,
                            "Date": dt,
                            "Month": month_num,
                            "Sales": pd.to_numeric(value, errors="coerce"),
                            "SourceSheet": sheet_name
                        })
                    except:
                        continue

        if all_data:
            result_df = pd.DataFrame(all_data)
            result_df = result_df.dropna(subset=["Sales"])
            return [result_df]
        else:
            return []

    except Exception as e:
        st.error(f"Error processing {sheet_name}: {e}")
        return []

def long_from_wide(df, area_col, state_col, year_col, month_cols):
    parts = []
    for _, row in df.iterrows():
        fy_start = to_fiscal_year_start(row[year_col]) if year_col else None
        # If year missing, we will not set date; will handle later
        for mcol in month_cols:
            val = row[mcol]
            # find month number for this column
            mkey = mcol.strip().lower().split()[-1]
            if mkey not in MONTHS_MAP:
                # try fallback: any token
                found = None
                for tok in mcol.strip().lower().split():
                    if tok in MONTHS_MAP:
                        found = MONTHS_MAP[tok]; break
                if not found:
                    continue
                mnum = found
            else:
                mnum = MONTHS_MAP[mkey]
            # Build calendar date; if fiscal year start is known: April maps to month 4 in fy_start.year
            if fy_start is not None:
                # months Apr..Dec of Y, Jan..Mar of Y+1
                if mnum >= 4:
                    y = fy_start.year
                else:
                    y = fy_start.year + 1
                dt = pd.Timestamp(year=y, month=mnum, day=1)
            else:
                # fallback to 1st of month in an arbitrary year, will sort later
                dt = pd.NaT
            parts.append({
                "Area": row[area_col] if area_col else "All",
                "State": row[state_col] if state_col else None,
                "FY": row[year_col] if year_col else None,
                "MonthName": mcol,
                "Date": dt,
                "Month": mnum,
                "Sales": pd.to_numeric(val, errors="coerce")
            })
    long = pd.DataFrame(parts)
    long = long.dropna(subset=["Sales"])
    # If Date missing, try to infer by ordering within FY
    if long["Date"].isna().any() and year_col:
        long["FY_start"] = long["FY"].apply(to_fiscal_year_start)
        def build_date(r):
            if pd.isna(r["Date"]) and pd.notna(r["FY_start"]):
                if r["Month"] >= 4:
                    y = r["FY_start"].year
                else:
                    y = r["FY_start"].year + 1
                return pd.Timestamp(year=y, month=int(r["Month"]), day=1)
            return r["Date"]
        long["Date"] = long.apply(build_date, axis=1)
        long = long.drop(columns=["FY_start"])
    return long

def fit_forecast(ts, horizon):
    # Handle degenerate series
    if ts.dropna().sum() == 0 or ts.dropna().nunique() <= 1:
        # naïve repeat or zeros
        last = ts.dropna().iloc[-1] if ts.dropna().size else 0.0
        start_date = ts.index[-1] if len(ts.index) > 0 else pd.Timestamp.now()
        fc = pd.Series([last]*horizon, index=pd.date_range(start_date+pd.offsets.MonthBegin(), periods=horizon, freq="MS"))
        return fc, None

    # Choose seasonal mode
    has_zero_or_neg = (ts <= 0).any()
    seasonal = 'add' if has_zero_or_neg else 'mul'
    try:
        model = ExponentialSmoothing(
            ts,
            trend='add',
            seasonal=seasonal,
            seasonal_periods=12,
            initialization_method="estimated"
        ).fit(optimized=True, use_brute=True)
        fc_index = pd.date_range(ts.index[-1] + pd.offsets.MonthBegin(), periods=horizon, freq="MS")
        forecast = pd.Series(model.forecast(horizon), index=fc_index)
        return forecast, model
    except Exception as e:
        # fallback: seasonal naive (last year same month)
        fc_index = pd.date_range(ts.index[-1] + pd.offsets.MonthBegin(), periods=horizon, freq="MS")
        forecast_vals = []
        for i, dt in enumerate(fc_index, start=1):
            same_month_last_year = dt - relativedelta(years=1)
            val = ts.get(same_month_last_year, np.nan)
            if np.isnan(val):
                val = ts.iloc[-12] if len(ts) >= 12 else ts.iloc[-1]
            forecast_vals.append(val)
        return pd.Series(forecast_vals, index=fc_index), None

def make_report(all_forecasts, hist_long, horizon=3, profit_margin=15.0):
    # Combine into one Excel-like buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Forecast Target For Areas (Enhanced Report)
        target_rows = []
        month_names = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"]

        for area, df_fc in all_forecasts.items():
            if len(df_fc) > 0:
                # Get the first date for this area
                start_date = df_fc.iloc[0]["Date"]

                # Create row with area and date
                row = {"Area": area, "Date": start_date}

                # Add target columns for each month in horizon
                for i in range(min(horizon, len(df_fc))):
                    forecast_value = df_fc.iloc[i]["Forecast"]
                    # Apply profit margin to get target
                    target_value = forecast_value * (1 + profit_margin / 100)

                    # Get month name for column header
                    month_date = df_fc.iloc[i]["Date"]
                    month_name = month_names[month_date.month - 1]
                    column_name = f"{month_name} Target"

                    row[column_name] = round(target_value, 2)

                target_rows.append(row)

        if target_rows:
            targets_df = pd.DataFrame(target_rows).sort_values("Area")
            targets_df.to_excel(writer, sheet_name="Forecast Target For Areas", index=False)

        # Summary
        summary_rows = []
        for area, df_fc in all_forecasts.items():
            total = df_fc["Forecast"].sum()
            summary_rows.append({"Area": area, "Forecast_Total": float(total)})
        pd.DataFrame(summary_rows).sort_values("Area").to_excel(writer, sheet_name="Summary", index=False)

        # Per-area forecast
        combined = []
        for area, df_fc in all_forecasts.items():
            tmp = df_fc.copy()
            tmp.insert(0, "Area", area)
            combined.append(tmp)
        if combined:
            pd.concat(combined, ignore_index=True).to_excel(writer, sheet_name="Area_Forecast", index=False)

        # Historical detail
        hist_long.to_excel(writer, sheet_name="Historical_Long", index=False)

    output.seek(0)
    return output

# ----------------------
# UI
# ----------------------
st.title("🌴 Kerala Territory Sales Forecast & Targets")

st.markdown("""
Upload your sales Excel and get **Kerala territory-wise monthly forecasts** for future months (Aug 2025 onwards).
- Focuses on **19 Kerala territories** from Route Sales data.
- Processes **yearly sheets (2018-2026)** with territory-wise historical data.
- Generates **monthly sales targets** for Kerala area managers.
- Models each Territory with **Holt-Winters (Exponential Smoothing)**.
- Provides **Aug 2025 - Mar 2026** projections based on historical patterns.
""")

uploaded = st.file_uploader("Upload Excel (.xlsx, .xls)", type=["xlsx", "xls"])

# Sidebar controls
default_h = 3
horizon = st.sidebar.number_input("Forecast horizon (months)", min_value=1, max_value=24, value=default_h, step=1)
target_only_next = st.sidebar.checkbox("Only next month target", value=True)
profit_margin = st.sidebar.number_input("Profit Margin (%)", min_value=0.0, max_value=100.0, value=15.0, step=0.5, help="Profit margin to add to forecast for target calculation")

st.sidebar.markdown("---")

# Kerala territories dropdown
kerala_territories_list = [
    'TRIVANDRUM', 'NEYYATINKARA', 'KOLLAM', 'PATHANAMTHITTA', 'KOTTAYAM',
    'ALAPPUZHA', 'IDUKKI', 'MOOVATTUPUZHA', 'ERNAMKULAM', 'PALAKKAD',
    'THRISSUR', 'EDAPAL', 'MALAPPURAM', 'KOZHIKODE CITY', 'VADAKARA',
    'WAYANAD', 'THALASSERY', 'KANNUR', 'KASARGOD'
]

selected_territory = st.sidebar.selectbox(
    "Select Territory for Detailed View",
    options=['All Territories'] + kerala_territories_list,
    index=0
)

aggregate_level = st.sidebar.selectbox("Aggregation level", ["Area", "State+Area", "All Areas"])

st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Excel Report Features")
st.sidebar.markdown(f"• **Horizon**: {horizon} month(s)")
st.sidebar.markdown(f"• **Profit Margin**: {profit_margin}%")
st.sidebar.markdown("• **Target Calculation**: Forecast + Profit Margin")
st.sidebar.markdown("• **Dynamic Columns**: Based on selected horizon")

# Sales Comparison Dashboard
if uploaded is not None:
    st.markdown("---")
    st.markdown("## 📊 Sales Comparison Dashboard")

    # Dashboard controls
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        sales_type = st.selectbox(
            "Sales Type",
            options=['All', 'Route Sales', 'MSD Sales', 'Inter Unit Sales'],
            index=0
        )

    with col2:
        report_period = st.selectbox(
            "Report Period",
            options=['Yearly', 'Monthly', 'Quarterly'],
            index=0
        )

    with col3:
        currency_format = st.selectbox(
            "Currency Format",
            options=['Lakhs (₹L)', 'Millions (₹M)'],
            index=0
        )

    with col4:
        chart_type = st.selectbox(
            "Chart Type",
            options=['Bar Chart', 'Line Chart', 'Area Chart'],
            index=0
        )

    # Always show the dashboard with exact data (no need to process from Excel)
    chart_data = prepare_chart_data(sales_type, currency_format)

    # Create dashboard table in the requested format
    st.markdown("### 📊 Sales Dashboard")

    # Create the dashboard table
    if sales_type == 'All':
        # Create comprehensive dashboard table with exact data
        dashboard_table = create_dashboard_table(currency_format)
        if dashboard_table is not None:
            st.dataframe(dashboard_table, use_container_width=True)

    if not chart_data.empty:
        # Set value suffix based on currency format
        value_suffix = 'L' if currency_format == 'Lakhs (₹L)' else 'M'

        # Create dashboard visualization
        col1, col2 = st.columns([2, 1])

        with col1:
            # Create chart based on selection
            if sales_type == 'All':
                # For "All" option, show stacked chart by sales category
                if chart_type == 'Bar Chart':
                    fig = px.bar(
                        chart_data,
                        x='Year',
                        y='Value',
                        color='Sales_Category',
                        title=f'All Sales Types - {report_period} Report',
                        labels={'Value': f'Sales (₹{value_suffix})', 'Year': 'Fiscal Year'},
                        barmode='stack'
                    )
                elif chart_type == 'Line Chart':
                    fig = px.line(
                        chart_data,
                        x='Year',
                        y='Value',
                        color='Sales_Category',
                        title=f'All Sales Types - {report_period} Report',
                        labels={'Value': f'Sales (₹{value_suffix})', 'Year': 'Fiscal Year'}
                    )
                else:  # Area Chart
                    fig = px.area(
                        chart_data,
                        x='Year',
                        y='Value',
                        color='Sales_Category',
                        title=f'All Sales Types - {report_period} Report',
                        labels={'Value': f'Sales (₹{value_suffix})', 'Year': 'Fiscal Year'}
                    )
            else:
                # For individual sales types
                if chart_type == 'Bar Chart':
                    fig = px.bar(
                        chart_data,
                        x='Year',
                        y='Value',
                        title=f'{sales_type} - {report_period} Report',
                        labels={'Value': f'Sales (₹{value_suffix})', 'Year': 'Fiscal Year'}
                    )
                elif chart_type == 'Line Chart':
                    fig = px.line(
                        chart_data,
                        x='Year',
                        y='Value',
                        title=f'{sales_type} - {report_period} Report',
                        labels={'Value': f'Sales (₹{value_suffix})', 'Year': 'Fiscal Year'}
                    )
                else:  # Area Chart
                    fig = px.area(
                        chart_data,
                        x='Year',
                        y='Value',
                        title=f'{sales_type} - {report_period} Report',
                        labels={'Value': f'Sales (₹{value_suffix})', 'Year': 'Fiscal Year'}
                    )

            fig.update_layout(
                height=400,
                yaxis=dict(
                    tickformat=".0f",
                    ticksuffix=value_suffix
                )
            )
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            # Summary metrics
            st.markdown("#### 📈 Summary")

            # Custom CSS for distinct metric boxes
            st.markdown("""
            <style>
            div[data-testid="metric-container"] {
                background-color: #ffffff;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                padding: 12px;
                margin: 4px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                transition: box-shadow 0.2s ease;
            }
            div[data-testid="metric-container"]:hover {
                box-shadow: 0 4px 8px rgba(0,0,0,0.15);
                border-color: #adb5bd;
            }
            div[data-testid="metric-container"] label {
                font-size: 11px !important;
                font-weight: 600 !important;
                color: #6c757d !important;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            div[data-testid="metric-container"] div[data-testid="metric-value"] {
                font-size: 16px !important;
                font-weight: 700 !important;
                color: #212529 !important;
                margin-top: 4px !important;
            }
            div[data-testid="metric-container"] div {
                margin: 0px !important;
                padding: 0px !important;
            }
            </style>
            """, unsafe_allow_html=True)

            if sales_type == 'All':
                total_sales = chart_data['Value'].sum()
                avg_sales = chart_data.groupby('Year')['Value'].sum().mean()
                years_count = chart_data['Year'].nunique()

                # Main metrics in 2 rows for better readability
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total Sales", f"₹{total_sales:,.1f}{value_suffix}")
                    st.metric("Years Covered", f"{years_count}")
                with col2:
                    st.metric("Average/Year", f"₹{avg_sales:,.1f}{value_suffix}")
                    # Top performers
                    if len(chart_data) > 0:
                        top_year = chart_data.groupby('Year')['Value'].sum().idxmax()
                        st.metric("Best Year", str(top_year))

                # Category breakdown
                st.markdown("---")
                st.markdown("#### By Category")
                category_totals = chart_data.groupby('Sales_Category')['Value'].sum().sort_values(ascending=False)

                # Display categories in a simple layout
                for category, total in category_totals.items():
                    st.metric(f"{category}", f"₹{total:,.1f}{value_suffix}")

            else:
                total_sales = chart_data['Value'].sum()
                avg_sales = chart_data['Value'].mean()
                years_count = len(chart_data)

                # Main metrics in 2 columns for better readability
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total Sales", f"₹{total_sales:,.1f}{value_suffix}")
                    st.metric("Years Covered", f"{years_count}")
                with col2:
                    st.metric("Average/Year", f"₹{avg_sales:,.1f}{value_suffix}")
                    # Top performers
                    if len(chart_data) > 0:
                        top_year = chart_data.loc[chart_data['Value'].idxmax(), 'Year']
                        st.metric("Best Year", str(top_year))

st.markdown("---")
if target_only_next:
    horizon = 1

if uploaded is not None:
    # Cache file bytes for better performance
    file_bytes = uploaded.getvalue()

    # Show progress indicator
    with st.spinner('Processing Excel file...'):
        try:
            raw = read_excel_cached(file_bytes, sheet_name=None, engine="openpyxl")
        except Exception:
            raw = read_excel_cached(file_bytes, sheet_name=None)  # fallback for .xls

    all_long = []

    # Filter out unwanted sheets
    excluded_sheets = ['2017-2018', '2018-2019 (Old)', 'Sheet1', 'Sheet2', 'Sheet3']
    relevant_sheets = {k: v for k, v in raw.items() if k not in excluded_sheets}



    for sheet_name, df in relevant_sheets.items():
        if df is None or df.empty:
            continue

        # Special handling for Monthly Comparison sheets
        if 'MONTHLY COMPARISON' in sheet_name.upper():
            processed_data = process_monthly_comparison_sheet(uploaded, sheet_name)
            if processed_data:
                all_long.extend(processed_data)
            continue

        # Special handling for Comparison Report sheet
        if 'COMPARISON REPORT' in sheet_name.upper():
            processed_data = process_comparison_report_sheet(uploaded, sheet_name)
            if processed_data:
                all_long.extend(processed_data)
            continue

        # Special handling for yearly sheets with territory data
        if any(char.isdigit() for char in sheet_name) and '-' in sheet_name:
            processed_data = process_territory_data_from_yearly_sheets(uploaded, sheet_name)
            if processed_data:
                all_long.extend(processed_data)
            continue

        # Try to find the actual data table by looking for month patterns in rows
        df_processed = None

        # For this specific Excel structure, try row 6 first (where months are located)
        target_rows = [6, 8, 5, 7, 9, 10]  # Based on our analysis

        for header_row in target_rows:
            if header_row >= len(df):
                continue

            try:
                # Read with specific header row
                if uploaded.name.endswith('.xls'):
                    df_test = pd.read_excel(uploaded, sheet_name=sheet_name, header=header_row)
                else:
                    df_test = pd.read_excel(uploaded, sheet_name=sheet_name, header=header_row, engine="openpyxl")

                # Check if we have month columns in the actual data
                month_cols_found = []
                expected_months = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
                for col in df_test.columns:
                    if col in expected_months:
                        month_cols_found.append(col)

                if len(month_cols_found) >= 8:  # Found at least 8 month columns
                    df_processed = df_test
                    break

            except Exception as e:
                continue

        # If still not found, try the advanced scanning approach
        if df_processed is None:
            # Look for month names in the actual data cells
            for row_idx in range(min(20, len(df))):
                row_data = df.iloc[row_idx].astype(str).str.lower()
                month_count = 0
                for cell in row_data:
                    if any(month in cell for month in ["april", "may", "june", "july", "august", "september", "october", "november", "december", "january", "february", "march"]):
                        month_count += 1

                if month_count >= 8:  # Found a row with multiple month names
                    try:
                        # Use this row as header
                        if uploaded.name.endswith('.xls'):
                            df_test = pd.read_excel(uploaded, sheet_name=sheet_name, header=row_idx)
                        else:
                            df_test = pd.read_excel(uploaded, sheet_name=sheet_name, header=row_idx, engine="openpyxl")

                        # Verify month columns
                        month_cols_found = []
                        expected_months = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
                        for col in df_test.columns:
                            if col in expected_months:
                                month_cols_found.append(col)

                        if len(month_cols_found) >= 8:
                            df_processed = df_test
                            break
                    except:
                        continue

        if df_processed is None:
            continue

        df = df_processed.copy()
        month_cols = guess_month_cols(df)
        if not month_cols:
            continue
        area_col, state_col, year_col = detect_keys(df)

        # If area not present, try to extract from sheet name or create meaningful areas
        if area_col is None:
            # For yearly sheets, extract area information from the data structure
            if any(char.isdigit() for char in sheet_name):
                # Look for area information in the data itself
                area_found = False
                for col in df.columns:
                    if 'particular' in str(col).lower():
                        # Check if this column contains area-like data
                        sample_data = df[col].dropna().astype(str)
                        area_keywords = ['KERALA', 'KARNATAKA', 'TAMIL', 'ZONE', 'REGION', 'STATE', 'MSD']
                        for val in sample_data.head(20):
                            if any(keyword in val.upper() for keyword in area_keywords):
                                df["Area"] = val.upper()
                                area_col = "Area"
                                area_found = True
                                break
                        if area_found:
                            break

                if not area_found:
                    # Create area from sheet name
                    df["Area"] = f"FY_{sheet_name.replace('-', '_')}"
                    area_col = "Area"

        # Keep only relevant cols
        keep = [c for c in [area_col, state_col, year_col] if c is not None] + month_cols
        keep = [c for c in keep if c in df.columns]
        if not keep:
            continue
        dfx = df[keep].copy()
        long = long_from_wide(dfx, area_col, state_col, year_col, month_cols)
        long["SourceSheet"] = sheet_name
        all_long.append(long)

    if not all_long:
        st.error("Could not detect any monthly data. Please ensure month columns (Apr..Mar) exist.")
        st.stop()

    long_all = pd.concat(all_long, ignore_index=True)
    # Clean
    long_all["Area"] = long_all["Area"].fillna("Unspecified")
    long_all["Date"] = pd.to_datetime(long_all["Date"])
    long_all = long_all.dropna(subset=["Date"])
    long_all = long_all.sort_values(["Area","Date"])
    # Remove completely null/zero sales rows
    # (we keep zeros; drop NaN handled already)

    # UI filters
    min_dt = long_all["Date"].min()
    max_dt = long_all["Date"].max()
    st.caption(f"Detected data range: {min_dt.date()} → {max_dt.date()}")

    # Area filter list
    areas = sorted(long_all["Area"].dropna().unique().tolist())

    # Default Kerala territories to be pre-selected
    default_kerala_territories = [
        'TRIVANDRUM', 'NEYYATINKARA', 'KOLLAM', 'PATHANAMTHITTA', 'KOTTAYAM',
        'ALAPPUZHA', 'IDUKKI', 'MOOVATTUPUZHA', 'ERNAMKULAM', 'PALAKKAD',
        'THRISSUR', 'EDAPAL', 'MALAPPURAM', 'KOZHIKODE CITY', 'VADAKARA',
        'WAYANAD', 'THALASSERY', 'KANNUR', 'KASARGOD'
    ]

    # Filter to only include territories that exist in the data
    default_selected = [area for area in default_kerala_territories if area in areas]

    # If no Kerala territories found, fall back to first 10 areas
    if not default_selected:
        default_selected = areas[: min(10, len(areas))]

    # Apply territory filter from sidebar if specific territory is selected
    if selected_territory != 'All Territories':
        if selected_territory in areas:
            selected_areas = [selected_territory]
        else:
            selected_areas = st.multiselect("Select Kerala Territories", areas, default=default_selected)
    else:
        selected_areas = st.multiselect("Select Kerala Territories", areas, default=default_selected)

    # Aggregate level transform
    if aggregate_level == "All Areas":
        long_all["GroupKey"] = "All"
    elif aggregate_level == "State+Area" and "State" in long_all.columns and long_all["State"].notna().any():
        long_all["GroupKey"] = long_all["State"].fillna("NA") + " - " + long_all["Area"].fillna("NA")
    else:
        long_all["GroupKey"] = long_all["Area"].fillna("Unspecified")

    # Filter to selected areas when applicable
    if aggregate_level != "All Areas":
        long_all = long_all[long_all["Area"].isin(selected_areas)]

    # Build monthly totals per group
    grp = long_all.groupby(["GroupKey","Date"], as_index=False)["Sales"].sum()

    # Convert sales to lakhs for display
    grp["Sales_Lakhs"] = grp["Sales"].apply(convert_to_lakhs)



    # Forecast per group
    all_fc = {}
    for g, gdf in grp.groupby("GroupKey"):
        ts = gdf.set_index("Date")["Sales_Lakhs"].asfreq("MS")
        # Fill missing months with 0 (or you may prefer NaN + interpolation)
        if ts.isna().any():
            # forward fill for short gaps else zeros
            ts = ts.ffill().fillna(0.0)
        fc, model = fit_forecast(ts, horizon)
        df_fc = pd.DataFrame({"Date": fc.index, "Forecast": fc.values})
        all_fc[g] = df_fc

    # Combine forecasts for display
    fc_disp = []
    for g, df_fc in all_fc.items():
        tmp = df_fc.copy()
        tmp.insert(0, "Group", g)
        fc_disp.append(tmp)
    fc_disp = pd.concat(fc_disp, ignore_index=True)



    # Downloadable Excel report
    report_buf = make_report(all_fc, long_all, horizon, profit_margin)
    st.download_button(
        label="⬇️ Download Forecast Report (Excel)",
        data=report_buf,
        file_name="area_monthly_forecast.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Target table for immediate month
    if target_only_next:
        st.subheader("🎯 Territory-wise Sales Targets for Area Managers")

        # Define the Kerala territories for reference
        kerala_territories_display = [
            'TRIVANDRUM', 'NEYYATINKARA', 'KOLLAM', 'PATHANAMTHITTA', 'KOTTAYAM',
            'ALAPPUZHA', 'IDUKKI', 'MOOVATTUPUZHA', 'ERNAMKULAM', 'PALAKKAD',
            'THRISSUR', 'EDAPAL', 'MALAPPURAM', 'KOZHIKODE CITY', 'VADAKARA',
            'WAYANAD', 'THALASSERY', 'KANNUR', 'KASARGOD'
        ]

        next_targets = []
        territory_targets = []

        for g, df_fc in all_fc.items():
            row = df_fc.iloc[0]
            target_value = float(row["Forecast"])

            # Check if this is one of our Kerala territories
            is_kerala_territory = g in kerala_territories_display

            target_data = {
                "Area/Region": g,
                "Target_Month": row["Date"].strftime("%b %Y"),
                "Sales_Target": format_lakhs(target_value),
                "Target_Value": target_value,
                "Target_Value_Lakhs": target_value,
                "Is_Kerala_Territory": is_kerala_territory
            }

            next_targets.append(target_data)
            if is_kerala_territory:
                territory_targets.append(target_data)

        targets_df = pd.DataFrame(next_targets)
        if len(targets_df) > 0:
            targets_df = targets_df.sort_values("Target_Value", ascending=False)

        territory_df = pd.DataFrame(territory_targets)
        if len(territory_df) > 0:
            territory_df = territory_df.sort_values("Target_Value", ascending=False)

        # Display Territory Targets Prominently
        if len(territory_df) > 0 and 'Target_Value' in territory_df.columns:
            st.markdown("### 🌴 **KERALA TERRITORY SALES TARGETS (Aug 2025 - Mar 2026)**")
            st.markdown(f"**{len(territory_df)} out of 19 Kerala territories found in data**")

            # Show territory targets in a prominent table
            territory_display = territory_df[["Area/Region", "Sales_Target"]].copy()
            territory_display.columns = ["Territory", "Next Month Target"]

            col1, col2 = st.columns([2, 1])

            with col1:
                st.dataframe(territory_display, use_container_width=True, hide_index=True)

            with col2:
                # Territory summary
                territory_total = territory_df["Target_Value"].sum()
                territory_avg = territory_df["Target_Value"].mean()
                top_territory = territory_df.iloc[0]["Area/Region"] if len(territory_df) > 0 else "N/A"

                st.metric("Total Territory Target", format_lakhs(territory_total))
                st.metric("Average per Territory", format_lakhs(territory_avg))
                st.metric("Top Territory", top_territory)

            # Territory chart
            st.markdown("### 📊 Territory Target Distribution")
            fig_territory = px.bar(territory_df, x="Area/Region", y="Target_Value_Lakhs",
                                 title="Sales Targets by Territory",
                                 labels={"Target_Value_Lakhs": "Sales Target (₹ Lakhs)", "Area/Region": "Territory"},
                                 color="Target_Value_Lakhs",
                                 color_continuous_scale="viridis")
            fig_territory.update_layout(
                xaxis_tickangle=-45,
                height=500,
                yaxis=dict(
                    tickformat=".0f",
                    ticksuffix="L"
                )
            )
            st.plotly_chart(fig_territory, use_container_width=True)

            # Missing Kerala territories alert
            found_territories = set(territory_df["Area/Region"].tolist())
            missing_territories = set(kerala_territories_display) - found_territories

        # Display all areas (including non-territories)
        st.markdown("### 📈 All Areas/Regions")
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### All Area Targets")
            display_df = targets_df[["Area/Region", "Sales_Target"]].copy()
            st.dataframe(display_df, use_container_width=True, hide_index=True)

        with col2:
            st.markdown("#### Target Distribution")
            # Add lakhs column for all targets
            targets_df['Target_Value_Lakhs'] = targets_df['Target_Value']

            fig_targets = px.bar(targets_df, x="Area/Region", y="Target_Value_Lakhs",
                               title="Sales Targets by All Areas",
                               labels={"Target_Value_Lakhs": "Sales Target (₹ Lakhs)", "Area/Region": "Area"})
            fig_targets.update_layout(
                xaxis_tickangle=-45,
                yaxis=dict(
                    tickformat=".0f",
                    ticksuffix="L"
                )
            )
            st.plotly_chart(fig_targets, use_container_width=True)

        # Summary statistics
        total_target = targets_df["Target_Value"].sum()
        avg_target = targets_df["Target_Value"].mean()
        top_area = targets_df.iloc[0]["Area/Region"]

        st.markdown("### 📋 Target Summary")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Target", format_lakhs(total_target))
        with col2:
            st.metric("Average Target", format_lakhs(avg_target))
        with col3:
            st.metric("Top Performing Area", top_area)

else:
    st.info("Upload your Excel file to begin.")
