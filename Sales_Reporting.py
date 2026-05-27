
# %%
from pathlib import Path
import pandas as pd
import numpy as np

# ==================================================================================================
# SETTINGS
# ==================================================================================================

BASE_FOLDER = Path(
    r"C:\Users\Natnael.Tesfagiorgis\OneDrive - Swinkels\Desktop\python\Backup Shared Folder"
)

E_FOLDER = Path(
    r"C:\Users\Natnael.Tesfagiorgis\OneDrive - Swinkels\Desktop\python\Backup Shared Folder"
)

HL_PER_CRATE = 0.0792
DEFAULT_TARGET_YEAR = 2026
SHOW_PREVIEW = False

# Final user-facing/reporting output rule:
# - Only management_kpi_table is exported for dashboard/reporting.
# - Other displayed tables are checks/diagnostics only.
# - final_kpi_report remains an internal backend calculation table only.


# ==================================================================================================
# FILE PATHS
# ==================================================================================================

files = {
    # Volume Plan
    "aop_target": BASE_FOLDER / "Volume Plan" / "AOP Target.xlsx",
    "sales_target": BASE_FOLDER / "Volume Plan" / "Sales Target.xlsx",
    "snop_target": BASE_FOLDER / "Volume Plan" / "SNOP Target.xlsx",

    # Left Tables / MDM
    "agent_left": BASE_FOLDER / "Volume Performance" / "left Tables" / "Agent Left table.xlsx",
    "brand_left": BASE_FOLDER / "Volume Performance" / "left Tables" / "Brand Left table.xlsx",
    "calendar_mdm": BASE_FOLDER / "Volume Performance" / "left Tables" / "Calander MDM.xlsx",
    "hl_converter": BASE_FOLDER / "Volume Performance" / "left Tables" / "HL Converter MDM.xlsx",
    "plant_code": BASE_FOLDER / "Volume Performance" / "left Tables" / "Plant Code.xlsx",

    # Stock
    "stock": BASE_FOLDER / "Volume Performance" / "Stock" / "Stock.xlsx",

    # Shipment
    "py_shipment": BASE_FOLDER / "Volume Performance" / "Shipment" / "PY" / "PY Shipment.xlsx",
    "actual_shipment": BASE_FOLDER / "Volume Performance" / "Shipment" / "YTD" / "Actual Shipment.xlsx",

    # Depletion
    "py_depletion": BASE_FOLDER / "Volume Performance" / "Depletion" / "PY" / "PY Deplition.xlsx",
    "ytd_depletion": BASE_FOLDER / "Volume Performance" / "Depletion" / "YTD" / "YTD Depletion.xlsx",
}

sheets = {
    "aop_target": "AOP_2026",
    "sales_target": "Sales Target",
    "snop_target": "SNOP Target",

    "agent_left": "Agent key",
    "brand_left": "Brand Key",
    "calendar_mdm": "Calander_2026",
    "hl_converter": "HL_Converter",
    "plant_code": "Plant",

    "stock": "Stock",

    # These will fall back to first available sheet if Sheet1 does not exist
    "py_shipment": "Sheet1",
    "actual_shipment": "Sheet1",

    "py_depletion": "Sheet1",
    "ytd_depletion": "YTD_Consolidated_Deplition",
}

# ==================================================================================================
# HELPER FUNCTIONS
# ==================================================================================================

def clean_columns(df):
    df = df.copy()
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.replace("  ", " ", regex=False)
    )
    return df


def clean_text(value):
    if pd.isna(value):
        return np.nan

    return (
        str(value)
        .strip()
        .upper()
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("  ", " ")
    )


def to_number(series):
    return pd.to_numeric(
        series
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("-", "0", regex=False)
        .str.strip()
        .replace(["", "nan", "None", "NaN"], np.nan),
        errors="coerce"
    )


def find_col(df, possible_cols, table_name):
    for col in possible_cols:
        if col in df.columns:
            return col

    raise KeyError(
        f"No matching column found in {table_name}. "
        f"Tried: {possible_cols}. "
        f"Available columns: {list(df.columns)}"
    )


def clean_working_day_flag(series):
    return (
        series
        .astype(str)
        .str.strip()
        .str.upper()
        .isin(["YES", "Y", "TRUE", "1", "WORKING DAY", "WORKING"])
    )


def parse_target_month(series, default_year=2026):
    """
    Safe month parser for target files.
    Handles:
    - Month number: 1, 2, 3...
    - Month text: Jan, January
    - Date values: 01-Jan-2026, 2026-01-01
    - Excel serial dates
    """

    month_lookup = {
        "JAN": 1, "JANUARY": 1,
        "FEB": 2, "FEBRUARY": 2,
        "MAR": 3, "MARCH": 3,
        "APR": 4, "APRIL": 4,
        "MAY": 5,
        "JUN": 6, "JUNE": 6,
        "JUL": 7, "JULY": 7,
        "AUG": 8, "AUGUST": 8,
        "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
        "OCT": 10, "OCTOBER": 10,
        "NOV": 11, "NOVEMBER": 11,
        "DEC": 12, "DECEMBER": 12,
    }

    def parse_one(value):
        if pd.isna(value):
            return pd.NaT

        if isinstance(value, pd.Timestamp):
            return pd.Timestamp(value.year, value.month, 1)

        text = str(value).strip()

        if text == "":
            return pd.NaT

        text_upper = text.upper()

        if text_upper in month_lookup:
            return pd.Timestamp(default_year, month_lookup[text_upper], 1)

        num = pd.to_numeric(text, errors="coerce")

        if pd.notna(num):

            # Month number
            if 1 <= num <= 12:
                return pd.Timestamp(default_year, int(num), 1)

            # YYYYMM format
            if 190001 <= num <= 209912:
                num_text = str(int(num))
                year = int(num_text[:4])
                month = int(num_text[4:6])

                if 1 <= month <= 12:
                    return pd.Timestamp(year, month, 1)

            # Excel serial date
            if 20000 <= num <= 60000:
                dt = pd.to_datetime(
                    num,
                    unit="D",
                    origin="1899-12-30",
                    errors="coerce"
                )

                if pd.notna(dt):
                    return pd.Timestamp(dt.year, dt.month, 1)

        dt = pd.to_datetime(text, errors="coerce", dayfirst=True)

        if pd.notna(dt):
            return pd.Timestamp(dt.year, dt.month, 1)

        dt = pd.to_datetime(f"{text} {default_year}", errors="coerce")

        if pd.notna(dt):
            return pd.Timestamp(dt.year, dt.month, 1)

        return pd.NaT

    return series.apply(parse_one)


# ==================================================================================================
# CHECK FILES EXIST
# ==================================================================================================

for name, path in files.items():
    if path.exists():
        print(f"✅ {name}: Found")
    else:
        print(f"❌ {name}: Missing -> {path}")

# ==================================================================================================
# READ ALL EXCEL FILES USING SPECIFIC SHEETS
# If configured sheet does not exist, use first available sheet.
# ==================================================================================================

tables = {}

for name, path in files.items():
    try:
        xls = pd.ExcelFile(path)
        available_sheets = xls.sheet_names
        requested_sheet = sheets.get(name)

        if requested_sheet in available_sheets:
            sheet_to_read = requested_sheet
        else:
            sheet_to_read = available_sheets[0]
            print(
                f"⚠️ {name}: Sheet '{requested_sheet}' not found. "
                f"Available sheets: {available_sheets}. "
                f"Using first sheet: '{sheet_to_read}'"
            )

        df = pd.read_excel(path, sheet_name=sheet_to_read)
        df = clean_columns(df)
        df = df.dropna(how="all").copy()

        tables[name] = df

        print(f"✅ Loaded {name}: {df.shape[0]:,} rows x {df.shape[1]:,} columns | Sheet: {sheet_to_read}")

    except Exception as e:
        print(f"❌ Failed to load {name}: {e}")


required_tables = [
    "aop_target",
    "sales_target",
    "snop_target",
    "agent_left",
    "brand_left",
    "calendar_mdm",
    "hl_converter",
    "plant_code",
    "stock",
    "py_shipment",
    "actual_shipment",
    "py_depletion",
    "ytd_depletion",
]

missing_loaded_tables = [
    name for name in required_tables
    if name not in tables
]

if missing_loaded_tables:
    raise KeyError(
        f"These tables failed to load: {missing_loaded_tables}. "
        f"Check file paths or sheet names."
    )

# ==================================================================================================
# ASSIGN TABLES
# ==================================================================================================

aop_target = tables["aop_target"]
sales_target = tables["sales_target"]
snop_target = tables["snop_target"]

agent_left = tables["agent_left"]
brand_left = tables["brand_left"]
calendar_mdm = tables["calendar_mdm"]
hl_converter = tables["hl_converter"]
plant_code = tables["plant_code"]

stock = tables["stock"]

py_shipment = tables["py_shipment"]
actual_shipment = tables["actual_shipment"]

py_depletion = tables["py_depletion"]
ytd_depletion = tables["ytd_depletion"]

# ==================================================================================================
# OPTIONAL PREVIEW
# ==================================================================================================

if SHOW_PREVIEW:
    for name, df in tables.items():
        print("=" * 100)
        print(f"TABLE: {name}")
        display(df.head(5))

# ==================================================================================================
# DIMENSIONS AND FACTS
# ==================================================================================================

dim_calendar = calendar_mdm.copy()
dim_agent = agent_left.copy()
dim_brand_raw = brand_left.copy()
dim_hl_converter = hl_converter.copy()
dim_plant = plant_code.copy()

fact_actual_shipment = actual_shipment.copy()
fact_py_shipment = py_shipment.copy()
fact_ytd_depletion = ytd_depletion.copy()
fact_py_depletion = py_depletion.copy()
fact_stock = stock.copy()

fact_aop_target = aop_target.copy()
fact_sales_target = sales_target.copy()
fact_snop_target = snop_target.copy()

fact_actual_shipment["Fact Source"] = "Actual Shipment"
fact_py_shipment["Fact Source"] = "PY Shipment"
fact_ytd_depletion["Fact Source"] = "YTD Depletion"
fact_py_depletion["Fact Source"] = "PY Depletion"
fact_stock["Fact Source"] = "Stock"
fact_aop_target["Fact Source"] = "AOP Target"
fact_sales_target["Fact Source"] = "Sales Target"
fact_snop_target["Fact Source"] = "SNOP Target"

# ==================================================================================================
# BRAND MDM
# Item Description = relationship key
# Report Naming = reporting name
# ==================================================================================================

dim_brand_reporting = dim_brand_raw[["Item Description", "Report Naming"]].copy()

dim_brand_reporting["Brand Key"] = dim_brand_reporting["Item Description"].apply(clean_text)
dim_brand_reporting["Reporting Name"] = dim_brand_reporting["Report Naming"].astype(str).str.strip()

dim_brand_reporting = (
    dim_brand_reporting
    .dropna(subset=["Brand Key"])
    .drop_duplicates(subset=["Brand Key"], keep="first")
    [["Brand Key", "Reporting Name"]]
)

# ==================================================================================================
# COMBINE TARGETS
# AOP is HL, converted to crates using 1 crate = 0.0792 HL.
# Sales Target and SNOP are already crates.
# ==================================================================================================

aop = fact_aop_target.copy()
sales = fact_sales_target.copy()
snop = fact_snop_target.copy()

sales_brand_col = find_col(
    sales,
    ["Brand", "Item Description", "Br"],
    "Sales Target"
)

sales_agent_col = find_col(
    sales,
    ["Sales Unit", "Agent"],
    "Sales Target"
)

sales_target_col = find_col(
    sales,
    ["Target", "Target Crates", "Target, Crates"],
    "Sales Target"
)

snop_brand_col = find_col(
    snop,
    ["Brand", "Item Description", "Br"],
    "SNOP Target"
)

snop_agent_col = find_col(
    snop,
    ["Sales Unit", "Agent"],
    "SNOP Target"
)

snop_target_col = find_col(
    snop,
    ["Target", "Target Crates", "Target, Crates"],
    "SNOP Target"
)

# --------------------------------------------------------------------------------------------------
# AOP TARGET
# --------------------------------------------------------------------------------------------------

fact_aop = pd.DataFrame({
    "Target Type": "AOP",
    "Month": aop["Month"],
    "DV": aop["Division"],
    "Agent": aop["Agent"],
    "Original Brand": aop["Item Description"],
    "Source UOM": "HL",
    "Target Value Source": to_number(aop["Target, HL"]),
})

fact_aop["Brand Key"] = fact_aop["Original Brand"].apply(clean_text)

fact_aop = fact_aop.merge(
    dim_brand_reporting,
    how="left",
    on="Brand Key"
)

fact_aop["Reporting Name"] = fact_aop["Reporting Name"].fillna(fact_aop["Original Brand"])

fact_aop["HL per Crate"] = HL_PER_CRATE
fact_aop["Target HL"] = fact_aop["Target Value Source"]
fact_aop["Target Crates"] = fact_aop["Target HL"] / HL_PER_CRATE

# --------------------------------------------------------------------------------------------------
# SALES TARGET
# --------------------------------------------------------------------------------------------------

fact_sales = pd.DataFrame({
    "Target Type": "Sales Target",
    "Month": sales["Month"],
    "DV": sales["Division"],
    "Agent": sales[sales_agent_col],
    "Original Brand": sales[sales_brand_col],
    "Source UOM": "Crates",
    "Target Value Source": to_number(sales[sales_target_col]),
})

fact_sales["Brand Key"] = fact_sales["Original Brand"].apply(clean_text)

fact_sales = fact_sales.merge(
    dim_brand_reporting,
    how="left",
    on="Brand Key"
)

fact_sales["Reporting Name"] = fact_sales["Reporting Name"].fillna(fact_sales["Original Brand"])

fact_sales["HL per Crate"] = HL_PER_CRATE
fact_sales["Target Crates"] = fact_sales["Target Value Source"]
fact_sales["Target HL"] = fact_sales["Target Crates"] * HL_PER_CRATE

# --------------------------------------------------------------------------------------------------
# SNOP TARGET
# --------------------------------------------------------------------------------------------------

fact_snop = pd.DataFrame({
    "Target Type": "SNOP",
    "Month": snop["Month"],
    "DV": snop["Division"],
    "Agent": snop[snop_agent_col],
    "Original Brand": snop[snop_brand_col],
    "Source UOM": "Crates",
    "Target Value Source": to_number(snop[snop_target_col]),
})

fact_snop["Brand Key"] = fact_snop["Original Brand"].apply(clean_text)

fact_snop = fact_snop.merge(
    dim_brand_reporting,
    how="left",
    on="Brand Key"
)

fact_snop["Reporting Name"] = fact_snop["Reporting Name"].fillna(fact_snop["Original Brand"])

fact_snop["HL per Crate"] = HL_PER_CRATE
fact_snop["Target Crates"] = fact_snop["Target Value Source"]
fact_snop["Target HL"] = fact_snop["Target Crates"] * HL_PER_CRATE

# --------------------------------------------------------------------------------------------------
# COMBINE TARGETS
# --------------------------------------------------------------------------------------------------

fact_target = pd.concat(
    [
        fact_aop,
        fact_sales,
        fact_snop
    ],
    ignore_index=True
)

fact_target = fact_target.dropna(how="all").copy()

valid_brand_keys = set(dim_brand_reporting["Brand Key"].dropna())

fact_target["Brand MDM Match"] = np.where(
    fact_target["Brand Key"].isin(valid_brand_keys),
    "Matched",
    "Not Matched"
)

fact_target = fact_target[
    [
        "Target Type",
        "Month",
        "DV",
        "Agent",
        "Original Brand",
        "Brand Key",
        "Reporting Name",
        "Brand MDM Match",
        "Source UOM",
        "Target Value Source",
        "HL per Crate",
        "Target HL",
        "Target Crates"
    ]
]

# ==================================================================================================
# UNMATCHED TARGET BRANDS
# ==================================================================================================

unmatched_target_brands = (
    fact_target[fact_target["Brand MDM Match"] == "Not Matched"]
    [["Target Type", "Original Brand", "Brand Key"]]
    .drop_duplicates()
    .sort_values(["Target Type", "Original Brand"])
)

print("=" * 120)
print("UNMATCHED TARGET BRANDS")
print("=" * 120)
display(unmatched_target_brands)

# ==================================================================================================
# FULL TARGET CHECK
# ==================================================================================================

target_check = (
    fact_target
    .groupby("Target Type", dropna=False)
    .agg(
        Total_Target_HL=("Target HL", "sum"),
        Total_Target_Crates=("Target Crates", "sum")
    )
    .reset_index()
)

target_check_display = target_check.copy()

target_check_display["Total_Target_HL"] = (
    target_check_display["Total_Target_HL"]
    .round(3)
    .map("{:,.3f}".format)
)

target_check_display["Total_Target_Crates"] = (
    target_check_display["Total_Target_Crates"]
    .round(0)
    .map("{:,.0f}".format)
)

print("=" * 120)
print("FULL TARGET CHECK")
print("=" * 120)
display(target_check_display)

display(fact_target.head(10))

# ==================================================================================================
# TARGET TO DATE BASED ON LAST ACTUAL SHIPMENT DATE
# Past months = 100%
# Current month = elapsed working days / total working days
# Future months = 0%
# ==================================================================================================

shipment_date_col = find_col(
    fact_actual_shipment,
    ["Billing Date", "Date", "Shipment Date", "Posting Date", "Invoice Date"],
    "fact_actual_shipment"
)

fact_actual_shipment[shipment_date_col] = pd.to_datetime(
    fact_actual_shipment[shipment_date_col],
    errors="coerce",
    dayfirst=True
)

LAST_SHIPMENT_DATE = fact_actual_shipment[shipment_date_col].max()

if pd.isna(LAST_SHIPMENT_DATE):
    raise ValueError("Last shipment date is blank. Check actual shipment date column.")

LAST_SHIPMENT_MONTH = pd.Timestamp(
    LAST_SHIPMENT_DATE.year,
    LAST_SHIPMENT_DATE.month,
    1
)

print("=" * 120)
print("LAST SHIPMENT DATE CHECK")
print("=" * 120)
print("Shipment date column used:", shipment_date_col)
print("Last shipment date:", LAST_SHIPMENT_DATE.date())
print("Last shipment month:", LAST_SHIPMENT_MONTH.date())

calendar_date_col = find_col(
    dim_calendar,
    ["Date", "Calendar Date", "Day", "GC Date", "Gregorian Date"],
    "dim_calendar"
)

working_day_col = find_col(
    dim_calendar,
    ["Working Day", "Working_Day", "Working day", "Is Working Day", "IsWorkingDay", "Workday", "Working"],
    "dim_calendar"
)

calendar_work = dim_calendar.copy()

calendar_work[calendar_date_col] = pd.to_datetime(
    calendar_work[calendar_date_col],
    errors="coerce",
    dayfirst=True
)

calendar_work["Is Working Day"] = clean_working_day_flag(calendar_work[working_day_col])

calendar_work = calendar_work[
    calendar_work[calendar_date_col].notna()
].copy()

calendar_work["Month Start"] = calendar_work[calendar_date_col].apply(
    lambda x: pd.Timestamp(x.year, x.month, 1)
)

calendar_work["Is Elapsed Working Day"] = (
    (calendar_work[calendar_date_col] <= LAST_SHIPMENT_DATE) &
    (calendar_work["Is Working Day"])
)

month_working_days = (
    calendar_work
    .groupby("Month Start", dropna=False)
    .agg(
        Month_Working_Days=("Is Working Day", "sum"),
        Elapsed_Working_Days=("Is Elapsed Working Day", "sum")
    )
    .reset_index()
)

month_working_days["Target_Progress_Ratio"] = np.select(
    [
        month_working_days["Month Start"] < LAST_SHIPMENT_MONTH,
        month_working_days["Month Start"] == LAST_SHIPMENT_MONTH,
        month_working_days["Month Start"] > LAST_SHIPMENT_MONTH
    ],
    [
        1,
        np.where(
            month_working_days["Month_Working_Days"] == 0,
            0,
            month_working_days["Elapsed_Working_Days"] / month_working_days["Month_Working_Days"]
        ),
        0
    ],
    default=0
)

month_working_days["Target_Progress_Ratio"] = (
    month_working_days["Target_Progress_Ratio"]
    .clip(0, 1)
)

print("=" * 120)
print("MONTH WORKING DAY PROGRESS CHECK")
print("=" * 120)
display(month_working_days)

# --------------------------------------------------------------------------------------------------
# APPLY TARGET TO DATE
# --------------------------------------------------------------------------------------------------

fact_target = fact_target.drop(
    columns=[
        "Month Start",
        "Month_Working_Days",
        "Elapsed_Working_Days",
        "Target_Progress_Ratio",
        "Target HL To Date",
        "Target Crates To Date"
    ],
    errors="ignore"
)

fact_target["Month Start"] = parse_target_month(
    fact_target["Month"],
    default_year=DEFAULT_TARGET_YEAR
)

print("=" * 120)
print("TARGET MONTH PARSE CHECK")
print("=" * 120)
display(
    fact_target[["Month", "Month Start"]]
    .drop_duplicates()
    .sort_values("Month Start")
)

fact_target = fact_target.merge(
    month_working_days[
        [
            "Month Start",
            "Month_Working_Days",
            "Elapsed_Working_Days",
            "Target_Progress_Ratio"
        ]
    ],
    how="left",
    on="Month Start"
)

fact_target["Target_Progress_Ratio"] = fact_target["Target_Progress_Ratio"].fillna(0)

fact_target["Target HL To Date"] = (
    fact_target["Target HL"] * fact_target["Target_Progress_Ratio"]
)

fact_target["Target Crates To Date"] = (
    fact_target["Target Crates"] * fact_target["Target_Progress_Ratio"]
)

# ==================================================================================================
# TARGET TO DATE CHECK
# ==================================================================================================

target_to_date_check = (
    fact_target
    .groupby("Target Type", dropna=False)
    .agg(
        Full_Target_HL=("Target HL", "sum"),
        Full_Target_Crates=("Target Crates", "sum"),
        Target_HL_To_Date=("Target HL To Date", "sum"),
        Target_Crates_To_Date=("Target Crates To Date", "sum")
    )
    .reset_index()
)

target_to_date_check_display = target_to_date_check.copy()

for col in [
    "Full_Target_HL",
    "Full_Target_Crates",
    "Target_HL_To_Date",
    "Target_Crates_To_Date"
]:
    target_to_date_check_display[col] = (
        target_to_date_check_display[col]
        .round(0)
        .map("{:,.0f}".format)
    )

print("=" * 120)
print("TARGET TO DATE CHECK")
print("=" * 120)
display(target_to_date_check_display)

# ==================================================================================================
# AOP MONTH CHECK
# ==================================================================================================

aop_month_check = (
    fact_target[fact_target["Target Type"] == "AOP"]
    .groupby(
        [
            "Month",
            "Month Start",
            "Month_Working_Days",
            "Elapsed_Working_Days",
            "Target_Progress_Ratio"
        ],
        dropna=False
    )
    .agg(
        Full_AOP_HL=("Target HL", "sum"),
        Full_AOP_Crates=("Target Crates", "sum"),
        AOP_HL_To_Date=("Target HL To Date", "sum"),
        AOP_Crates_To_Date=("Target Crates To Date", "sum")
    )
    .reset_index()
    .sort_values("Month Start")
)

aop_month_check_display = aop_month_check.copy()

for col in [
    "Full_AOP_HL",
    "Full_AOP_Crates",
    "AOP_HL_To_Date",
    "AOP_Crates_To_Date"
]:
    aop_month_check_display[col] = (
        aop_month_check_display[col]
        .round(0)
        .map("{:,.0f}".format)
    )

aop_month_check_display["Target_Progress_Ratio"] = (
    aop_month_check_display["Target_Progress_Ratio"]
    .map("{:.1%}".format)
)

print("=" * 120)
print("AOP MONTH TARGET TO DATE CHECK")
print("=" * 120)
display(aop_month_check_display)

# ==================================================================================================
# REPORT 1: REGIONAL TARGET REPORT TO DATE
# ==================================================================================================

regional_target_report = (
    fact_target
    .groupby(["DV", "Target Type"], dropna=False)
    .agg(
        Full_Target_Crates=("Target Crates", "sum"),
        Target_Crates_To_Date=("Target Crates To Date", "sum")
    )
    .reset_index()
)

regional_target_report_display = regional_target_report.copy()

for col in ["Full_Target_Crates", "Target_Crates_To_Date"]:
    regional_target_report_display[col] = (
        regional_target_report_display[col]
        .round(0)
        .map("{:,.0f}".format)
    )

print("=" * 120)
print("REGIONAL TARGET REPORT TO DATE")
print("=" * 120)
display(regional_target_report_display)

regional_target_pivot = (
    fact_target
    .pivot_table(
        index="DV",
        columns="Target Type",
        values="Target Crates To Date",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

for col in ["Sales Target", "SNOP", "AOP"]:
    if col not in regional_target_pivot.columns:
        regional_target_pivot[col] = 0

regional_target_pivot = regional_target_pivot[
    ["DV", "Sales Target", "SNOP", "AOP"]
]

regional_target_pivot_display = regional_target_pivot.copy()

for col in ["Sales Target", "SNOP", "AOP"]:
    regional_target_pivot_display[col] = (
        regional_target_pivot_display[col]
        .round(0)
        .map("{:,.0f}".format)
    )

print("=" * 120)
print("REGIONAL TARGET TO DATE PIVOT - CRATES")
print("=" * 120)
display(regional_target_pivot_display)

# ==================================================================================================
# SELECTED TARGET FOR FUTURE REPORTS
# Change this to: "Sales Target", "SNOP", or "AOP"
# ==================================================================================================

SELECTED_TARGET_TYPE = "Sales Target"

selected_target = fact_target[
    fact_target["Target Type"] == SELECTED_TARGET_TYPE
].copy()

selected_target_by_region = (
    selected_target
    .groupby("DV", dropna=False)
    .agg(
        Selected_Target_Crates=("Target Crates To Date", "sum"),
        Selected_Target_HL=("Target HL To Date", "sum")
    )
    .reset_index()
)

selected_target_by_region_display = selected_target_by_region.copy()

selected_target_by_region_display["Selected_Target_Crates"] = (
    selected_target_by_region_display["Selected_Target_Crates"]
    .round(0)
    .map("{:,.0f}".format)
)

selected_target_by_region_display["Selected_Target_HL"] = (
    selected_target_by_region_display["Selected_Target_HL"]
    .round(0)
    .map("{:,.0f}".format)
)

print("=" * 120)
print(f"SELECTED TARGET BY REGION: {SELECTED_TARGET_TYPE}")
print("=" * 120)
display(selected_target_by_region_display)

# ==================================================================================================
# STAR SCHEMA TABLE COLLECTION
# ==================================================================================================

dimensions = {
    "dim_calendar": dim_calendar,
    "dim_agent": dim_agent,
    "dim_brand_raw": dim_brand_raw,
    "dim_brand_reporting": dim_brand_reporting,
    "dim_hl_converter": dim_hl_converter,
    "dim_plant": dim_plant,
}

facts = {
    "fact_actual_shipment": fact_actual_shipment,
    "fact_py_shipment": fact_py_shipment,
    "fact_ytd_depletion": fact_ytd_depletion,
    "fact_py_depletion": fact_py_depletion,
    "fact_stock": fact_stock,
    "fact_target": fact_target,
    "selected_target": selected_target,
}
# %%
print("=" * 120)
print("ACTUAL SHIPMENT COLUMNS")
print("=" * 120)

print(list(fact_actual_shipment.columns))

display(fact_actual_shipment.head(10))
# %%


# %%
# ==================================================================================================
# REPORT 2: REGIONAL ACTUAL SHIPMENT VS SELECTED TARGET TO-DATE + VS PY
# Shipment files do NOT have DV. DV is mapped from Agent MDM.
#
# Shipment source columns confirmed:
#   Date
#   Item Description
#   Agent
#   Invoiced Quantity
#
# Agent MDM confirmed:
#   Name(S4H_C4C)
#   DV
# ==================================================================================================

SELECTED_TARGET_TYPE = "Sales Target"   # Options: "Sales Target", "SNOP", "AOP"

# ==================================================================================================
# 1. BUILD AGENT MAPPING TABLE
# New Agent MDM logic:
#   - Source agent name can come from S4H naming, depletion naming, stock naming, or target naming.
#   - Agent Reporting Name is the standard final display name used in all reports.
#   - After mapping, Agent Key is rebuilt from Agent Reporting Name so all agent sources consolidate.
# ==================================================================================================

agent_mdm = dim_agent.copy()
agent_mdm = clean_columns(agent_mdm)

agent_mdm_agent_col = find_col(
    agent_mdm,
    [
        "Name(S4H_C4C)",
        "Agent",
        "Sales Unit",
        "Sold-to Name",
        "Customer Name",
        "Customer",
        "Agent Name"
    ],
    "dim_agent"
)

agent_mdm_dv_col = find_col(
    agent_mdm,
    [
        "DV",
        "Division",
        "Region"
    ],
    "dim_agent"
)

agent_reporting_name_col = None
for col in [
    "Reporting Name",
    "Report Naming",
    "Agent Reporting Name",
    "Agent Report Naming",
    "Reporting_Name",
    "Reporting name",
    "Repoting Name",
    "Repoting name"
]:
    if col in agent_mdm.columns:
        agent_reporting_name_col = col
        break

if agent_reporting_name_col is None:
    print("⚠️ Agent reporting name column not found. Using source agent name as reporting name.")
    agent_mdm["Agent Reporting Name"] = agent_mdm[agent_mdm_agent_col]
    agent_reporting_name_col = "Agent Reporting Name"

dim_agent_reporting = agent_mdm[
    [
        agent_mdm_agent_col,
        agent_mdm_dv_col,
        agent_reporting_name_col
    ]
].copy()

dim_agent_reporting = dim_agent_reporting.rename(
    columns={
        agent_mdm_agent_col: "Agent Source Name",
        agent_mdm_dv_col: "DV_From_Agent_MDM",
        agent_reporting_name_col: "Agent Reporting Name"
    }
)

dim_agent_reporting["Agent Source Key"] = dim_agent_reporting["Agent Source Name"].apply(clean_text)
dim_agent_reporting["Agent Reporting Name"] = (
    dim_agent_reporting["Agent Reporting Name"]
    .fillna(dim_agent_reporting["Agent Source Name"])
    .astype(str)
    .str.strip()
)

dim_agent_reporting["Agent Reporting Key"] = dim_agent_reporting["Agent Reporting Name"].apply(clean_text)

dim_agent_reporting = (
    dim_agent_reporting
    .dropna(subset=["Agent Source Key"])
    .drop_duplicates(subset=["Agent Source Key"], keep="first")
    [
        [
            "Agent Source Key",
            "Agent Reporting Key",
            "Agent Reporting Name",
            "DV_From_Agent_MDM"
        ]
    ]
)

print("=" * 120)
print("AGENT MDM MAPPING CHECK - USING REPORTING NAME")
print("=" * 120)
print("Agent MDM source agent column:", agent_mdm_agent_col)
print("Agent MDM DV column:", agent_mdm_dv_col)
print("Agent MDM reporting name column:", agent_reporting_name_col)

display(dim_agent_reporting.head(10))


# ==================================================================================================
# 2. COLUMN MAPPING - ACTUAL SHIPMENT
# ==================================================================================================

actual_date_col = find_col(
    fact_actual_shipment,
    ["Date", "Billing Date", "Shipment Date", "Posting Date", "Invoice Date"],
    "fact_actual_shipment"
)

actual_agent_col = find_col(
    fact_actual_shipment,
    ["Agent", "Sales Unit", "Sold-to Name", "Customer Name", "Customer"],
    "fact_actual_shipment"
)

actual_brand_col = find_col(
    fact_actual_shipment,
    ["Item Description", "Brand", "Product", "Material Description"],
    "fact_actual_shipment"
)

actual_crates_col = find_col(
    fact_actual_shipment,
    ["Invoiced Quantity", "Shipment", "Shipment Crates", "Crates", "Quantity", "Qty", "Billed Quantity"],
    "fact_actual_shipment"
)


# ==================================================================================================
# 3. COLUMN MAPPING - PY SHIPMENT
# ==================================================================================================

py_date_col = find_col(
    fact_py_shipment,
    ["Date", "Billing Date", "Shipment Date", "Posting Date", "Invoice Date"],
    "fact_py_shipment"
)

py_agent_col = find_col(
    fact_py_shipment,
    ["Agent", "Sales Unit", "Sold-to Name", "Customer Name", "Customer"],
    "fact_py_shipment"
)

py_brand_col = find_col(
    fact_py_shipment,
    ["Item Description", "Brand", "Product", "Material Description"],
    "fact_py_shipment"
)

py_crates_col = find_col(
    fact_py_shipment,
    ["Invoiced Quantity", "Shipment", "Shipment Crates", "Crates", "Quantity", "Qty", "Billed Quantity"],
    "fact_py_shipment"
)

print("=" * 120)
print("SHIPMENT COLUMN MAPPING")
print("=" * 120)
print("Actual date column:", actual_date_col)
print("Actual agent column:", actual_agent_col)
print("Actual brand column:", actual_brand_col)
print("Actual crates column:", actual_crates_col)
print("-" * 120)
print("PY date column:", py_date_col)
print("PY agent column:", py_agent_col)
print("PY brand column:", py_brand_col)
print("PY crates column:", py_crates_col)


# ==================================================================================================
# 4. PREPARE ACTUAL SHIPMENT
# ==================================================================================================

actual_work = fact_actual_shipment.copy()

actual_work[actual_date_col] = pd.to_datetime(
    actual_work[actual_date_col],
    errors="coerce",
    dayfirst=True
)

actual_work = actual_work[
    actual_work[actual_date_col].notna()
].copy()

actual_work["Shipment Date"] = actual_work[actual_date_col]
actual_work["Agent"] = actual_work[actual_agent_col]
actual_work["Agent Key"] = actual_work["Agent"].apply(clean_text)

actual_work["Original Brand"] = actual_work[actual_brand_col]
actual_work["Brand Key"] = actual_work["Original Brand"].apply(clean_text)

actual_work["Actual Shipment Crates"] = to_number(actual_work[actual_crates_col])
actual_work["Actual Shipment HL"] = actual_work["Actual Shipment Crates"] * HL_PER_CRATE

# Map DV and standardized Agent Reporting Name from Agent MDM
actual_work = actual_work.merge(
    dim_agent_reporting,
    how="left",
    left_on="Agent Key",
    right_on="Agent Source Key"
)

actual_work["Agent Source Name"] = actual_work["Agent"]
actual_work["Agent"] = actual_work["Agent Reporting Name"].fillna(actual_work["Agent Source Name"])
actual_work["Agent Key"] = actual_work["Agent"].apply(clean_text)
actual_work["DV"] = actual_work["DV_From_Agent_MDM"]

# Map brand reporting name
actual_work = actual_work.merge(
    dim_brand_reporting,
    how="left",
    on="Brand Key"
)

actual_work["Reporting Name"] = actual_work["Reporting Name"].fillna(actual_work["Original Brand"])

actual_work["Brand MDM Match"] = np.where(
    actual_work["Brand Key"].isin(set(dim_brand_reporting["Brand Key"].dropna())),
    "Matched",
    "Not Matched"
)

actual_work["Agent MDM Match"] = np.where(
    actual_work["DV"].notna(),
    "Matched",
    "Not Matched"
)


# ==================================================================================================
# 5. APPLY CALENDAR WORKING DAY FILTER TO ACTUAL
# ==================================================================================================

calendar_for_actual = calendar_work[
    [
        calendar_date_col,
        "Is Working Day"
    ]
].copy()

calendar_for_actual = calendar_for_actual.rename(
    columns={
        calendar_date_col: "Shipment Date"
    }
)

actual_work = actual_work.merge(
    calendar_for_actual,
    how="left",
    on="Shipment Date"
)

actual_work["Is Working Day"] = actual_work["Is Working Day"].fillna(False)

actual_to_date = actual_work[
    actual_work["Shipment Date"] <= LAST_SHIPMENT_DATE
].copy()


# ==================================================================================================
# 6. PREPARE PY SHIPMENT
# Same period last year
# If actual cutoff = 2026-05-20, PY cutoff = 2025-05-20
# ==================================================================================================

PY_CUTOFF_DATE = LAST_SHIPMENT_DATE - pd.DateOffset(years=1)

py_work = fact_py_shipment.copy()

py_work[py_date_col] = pd.to_datetime(
    py_work[py_date_col],
    errors="coerce",
    dayfirst=True
)

py_work = py_work[
    py_work[py_date_col].notna()
].copy()

py_work["PY Shipment Date"] = py_work[py_date_col]
py_work["Agent"] = py_work[py_agent_col]
py_work["Agent Key"] = py_work["Agent"].apply(clean_text)

py_work["Original Brand"] = py_work[py_brand_col]
py_work["Brand Key"] = py_work["Original Brand"].apply(clean_text)

py_work["PY Shipment Crates"] = to_number(py_work[py_crates_col])
py_work["PY Shipment HL"] = py_work["PY Shipment Crates"] * HL_PER_CRATE

# Map DV and standardized Agent Reporting Name from Agent MDM
py_work = py_work.merge(
    dim_agent_reporting,
    how="left",
    left_on="Agent Key",
    right_on="Agent Source Key"
)

py_work["Agent Source Name"] = py_work["Agent"]
py_work["Agent"] = py_work["Agent Reporting Name"].fillna(py_work["Agent Source Name"])
py_work["Agent Key"] = py_work["Agent"].apply(clean_text)
py_work["DV"] = py_work["DV_From_Agent_MDM"]

# Map brand reporting name
py_work = py_work.merge(
    dim_brand_reporting,
    how="left",
    on="Brand Key"
)

py_work["Reporting Name"] = py_work["Reporting Name"].fillna(py_work["Original Brand"])

py_work["Brand MDM Match"] = np.where(
    py_work["Brand Key"].isin(set(dim_brand_reporting["Brand Key"].dropna())),
    "Matched",
    "Not Matched"
)

py_work["Agent MDM Match"] = np.where(
    py_work["DV"].notna(),
    "Matched",
    "Not Matched"
)


# ==================================================================================================
# 7. APPLY PY SAME-PERIOD FILTER
# Working-day logic maps PY date month/day onto current-year Calendar MDM.
# ==================================================================================================

def map_py_date_to_current_year(value, current_year):
    if pd.isna(value):
        return pd.NaT

    try:
        return pd.Timestamp(current_year, value.month, value.day)
    except ValueError:
        return pd.NaT


py_work["Calendar Match Date"] = py_work["PY Shipment Date"].apply(
    lambda x: map_py_date_to_current_year(x, LAST_SHIPMENT_DATE.year)
)

calendar_for_py = calendar_work[
    [
        calendar_date_col,
        "Is Working Day"
    ]
].copy()

calendar_for_py = calendar_for_py.rename(
    columns={
        calendar_date_col: "Calendar Match Date"
    }
)

py_work = py_work.merge(
    calendar_for_py,
    how="left",
    on="Calendar Match Date"
)

py_work["Is Working Day"] = py_work["Is Working Day"].fillna(False)

py_to_date = py_work[
    py_work["PY Shipment Date"] <= PY_CUTOFF_DATE
].copy()


# ==================================================================================================
# 8. REGIONAL ACTUAL, PY, AND TARGET
# ==================================================================================================

regional_actual = (
    actual_to_date
    .groupby("DV", dropna=False)
    .agg(
        Actual_Shipment_Crates=("Actual Shipment Crates", "sum"),
        Actual_Shipment_HL=("Actual Shipment HL", "sum")
    )
    .reset_index()
)

regional_py = (
    py_to_date
    .groupby("DV", dropna=False)
    .agg(
        PY_Shipment_Crates=("PY Shipment Crates", "sum"),
        PY_Shipment_HL=("PY Shipment HL", "sum")
    )
    .reset_index()
)

selected_target = fact_target[
    fact_target["Target Type"] == SELECTED_TARGET_TYPE
].copy()

regional_selected_target = (
    selected_target
    .groupby("DV", dropna=False)
    .agg(
        Target_Crates_To_Date=("Target Crates To Date", "sum"),
        Target_HL_To_Date=("Target HL To Date", "sum"),
        Full_Target_Crates=("Target Crates", "sum")
    )
    .reset_index()
)


# ==================================================================================================
# 9. MERGE REPORT
# ==================================================================================================

regional_actual_vs_target_py = (
    regional_selected_target
    .merge(regional_actual, how="left", on="DV")
    .merge(regional_py, how="left", on="DV")
)

for col in [
    "Actual_Shipment_Crates",
    "Actual_Shipment_HL",
    "PY_Shipment_Crates",
    "PY_Shipment_HL"
]:
    regional_actual_vs_target_py[col] = regional_actual_vs_target_py[col].fillna(0)

regional_actual_vs_target_py["Variance_vs_Target_Crates"] = (
    regional_actual_vs_target_py["Actual_Shipment_Crates"] -
    regional_actual_vs_target_py["Target_Crates_To_Date"]
)

regional_actual_vs_target_py["Achievement %"] = np.where(
    regional_actual_vs_target_py["Target_Crates_To_Date"] == 0,
    0,
    regional_actual_vs_target_py["Actual_Shipment_Crates"] /
    regional_actual_vs_target_py["Target_Crates_To_Date"]
)

regional_actual_vs_target_py["Variance_vs_PY_Crates"] = (
    regional_actual_vs_target_py["Actual_Shipment_Crates"] -
    regional_actual_vs_target_py["PY_Shipment_Crates"]
)

regional_actual_vs_target_py["Growth vs PY %"] = np.where(
    regional_actual_vs_target_py["PY_Shipment_Crates"] == 0,
    0,
    regional_actual_vs_target_py["Variance_vs_PY_Crates"] /
    regional_actual_vs_target_py["PY_Shipment_Crates"]
)

regional_actual_vs_target_py["Contribution %"] = np.where(
    regional_actual_vs_target_py["Actual_Shipment_Crates"].sum() == 0,
    0,
    regional_actual_vs_target_py["Actual_Shipment_Crates"] /
    regional_actual_vs_target_py["Actual_Shipment_Crates"].sum()
)

regional_actual_vs_target_py = regional_actual_vs_target_py.sort_values(
    "Actual_Shipment_Crates",
    ascending=False
)


# ==================================================================================================
# 10. DISPLAY REPORT
# ==================================================================================================

regional_actual_vs_target_py_display = regional_actual_vs_target_py.copy()

number_cols = [
    "Actual_Shipment_Crates",
    "Target_Crates_To_Date",
    "Variance_vs_Target_Crates",
    "PY_Shipment_Crates",
    "Variance_vs_PY_Crates",
    "Full_Target_Crates"
]

for col in number_cols:
    regional_actual_vs_target_py_display[col] = (
        regional_actual_vs_target_py_display[col]
        .round(0)
        .map("{:,.0f}".format)
    )

for col in ["Achievement %", "Growth vs PY %", "Contribution %"]:
    regional_actual_vs_target_py_display[col] = (
        regional_actual_vs_target_py_display[col]
        .map("{:.1%}".format)
    )

print("=" * 120)
print(f"REGIONAL ACTUAL SHIPMENT VS TARGET TO-DATE + VS PY: {SELECTED_TARGET_TYPE}")
print("=" * 120)
print("Actual cutoff date:", LAST_SHIPMENT_DATE.date())
print("PY cutoff date:", PY_CUTOFF_DATE.date())

display(
    regional_actual_vs_target_py_display[
        [
            "DV",
            "Actual_Shipment_Crates",
            "Target_Crates_To_Date",
            "Variance_vs_Target_Crates",
            "Achievement %",
            "PY_Shipment_Crates",
            "Variance_vs_PY_Crates",
            "Growth vs PY %",
            "Contribution %",
            "Full_Target_Crates"
        ]
    ]
)


# ==================================================================================================
# 11. DATA QUALITY CHECKS
# ==================================================================================================

unmatched_actual_agents = (
    actual_to_date[actual_to_date["Agent MDM Match"] == "Not Matched"]
    [[col for col in ["Agent Source Name", "Agent", "Agent Key"] if col in actual_to_date.columns]]
    .drop_duplicates()
    .sort_values("Agent")
)

unmatched_py_agents = (
    py_to_date[py_to_date["Agent MDM Match"] == "Not Matched"]
    [[col for col in ["Agent Source Name", "Agent", "Agent Key"] if col in py_to_date.columns]]
    .drop_duplicates()
    .sort_values("Agent")
)

unmatched_actual_brands = (
    actual_to_date[actual_to_date["Brand MDM Match"] == "Not Matched"]
    [["Original Brand", "Brand Key"]]
    .drop_duplicates()
    .sort_values("Original Brand")
)

unmatched_py_brands = (
    py_to_date[py_to_date["Brand MDM Match"] == "Not Matched"]
    [["Original Brand", "Brand Key"]]
    .drop_duplicates()
    .sort_values("Original Brand")
)

print("=" * 120)
print("UNMATCHED ACTUAL SHIPMENT AGENTS")
print("=" * 120)
display(unmatched_actual_agents)

print("=" * 120)
print("UNMATCHED PY SHIPMENT AGENTS")
print("=" * 120)
display(unmatched_py_agents)

print("=" * 120)
print("UNMATCHED ACTUAL SHIPMENT BRANDS")
print("=" * 120)
display(unmatched_actual_brands)

print("=" * 120)
print("UNMATCHED PY SHIPMENT BRANDS")
print("=" * 120)
display(unmatched_py_brands)


# ==================================================================================================
# 12. TOTAL CHECK
# ==================================================================================================

shipment_total_check = pd.DataFrame({
    "Metric": [
        "Actual Shipment Crates To Date",
        "Selected Target Crates To Date",
        "Variance vs Target",
        "PY Shipment Crates Same Period",
        "Variance vs PY"
    ],
    "Value": [
        actual_to_date["Actual Shipment Crates"].sum(),
        selected_target["Target Crates To Date"].sum(),
        actual_to_date["Actual Shipment Crates"].sum() - selected_target["Target Crates To Date"].sum(),
        py_to_date["PY Shipment Crates"].sum(),
        actual_to_date["Actual Shipment Crates"].sum() - py_to_date["PY Shipment Crates"].sum()
    ]
})

shipment_total_check["Value"] = (
    shipment_total_check["Value"]
    .round(0)
    .map("{:,.0f}".format)
)

print("=" * 120)
print("SHIPMENT TOTAL CHECK")
print("=" * 120)
display(shipment_total_check)
# %%







# %%


# %%


# ==================================================================================================
# %%
# ==================================================================================================
# REPORT 3: FINAL AGENT + BRAND KPI REPORT - OPTIMIZED
# Grain before UOM expansion:
#   Division + Agent + Brand
#
# Final output:
#   final_kpi_wide   = one row per Division + Agent + Brand in Crates only
#   final_kpi_report = two rows per Division + Agent + Brand: UOM = Crates and UOM = HL
#
# Design rules:
#   - Final base is the UNION of actual, PY, target, stock, and depletion keys.
#   - This keeps target-only agents, PY-only agents, shipment-only agents, stock-only agents,
#     and depletion-only agents in the report.
#   - Targets are shown side by side: AOP, SNOP, Sales.
#   - Actual shipment is not working-day filtered.
#   - Target to-date is already phased using Calendar MDM working days in fact_target.
#   - Stock Balance = Beginning Stock + MTD Shipment - MTD Depletion, clipped at zero.
#   - UOM slicer: use final_kpi_report and filter UOM to either Crates or HL.
# ==================================================================================================

# ==================================================================================================
# 1. DATE SETUP
# ==================================================================================================

CURRENT_MONTH_START = pd.Timestamp(LAST_SHIPMENT_DATE.year, LAST_SHIPMENT_DATE.month, 1)
CURRENT_MONTH_END = CURRENT_MONTH_START + pd.offsets.MonthEnd(0)

YEAR_START = pd.Timestamp(LAST_SHIPMENT_DATE.year, 1, 1)
YEAR_END = pd.Timestamp(LAST_SHIPMENT_DATE.year, 12, 31)

PY_CUTOFF_DATE = LAST_SHIPMENT_DATE - pd.DateOffset(years=1)
PY_MONTH_START = pd.Timestamp(PY_CUTOFF_DATE.year, PY_CUTOFF_DATE.month, 1)
PY_YEAR_START = pd.Timestamp(PY_CUTOFF_DATE.year, 1, 1)

remaining_working_days = calendar_work[
    (calendar_work[calendar_date_col] > LAST_SHIPMENT_DATE) &
    (calendar_work[calendar_date_col] <= CURRENT_MONTH_END) &
    (calendar_work["Is Working Day"])
].shape[0]

elapsed_working_days = calendar_work[
    (calendar_work[calendar_date_col] >= CURRENT_MONTH_START) &
    (calendar_work[calendar_date_col] <= LAST_SHIPMENT_DATE) &
    (calendar_work["Is Working Day"])
].shape[0]

elapsed_year_working_days = calendar_work[
    (calendar_work[calendar_date_col] >= YEAR_START) &
    (calendar_work[calendar_date_col] <= LAST_SHIPMENT_DATE) &
    (calendar_work["Is Working Day"])
].shape[0]

remaining_year_working_days = calendar_work[
    (calendar_work[calendar_date_col] > LAST_SHIPMENT_DATE) &
    (calendar_work[calendar_date_col] <= YEAR_END) &
    (calendar_work["Is Working Day"])
].shape[0]

elapsed_calendar_days = max((LAST_SHIPMENT_DATE - CURRENT_MONTH_START).days + 1, 1)
elapsed_working_days = max(elapsed_working_days, 1)
elapsed_year_working_days = max(elapsed_year_working_days, 1)

print("=" * 120)
print("FINAL KPI DATE CHECK")
print("=" * 120)
print("Current month start:", CURRENT_MONTH_START.date())
print("Last shipment date:", LAST_SHIPMENT_DATE.date())
print("Current month end:", CURRENT_MONTH_END.date())
print("PY month start:", PY_MONTH_START.date())
print("PY cutoff date:", PY_CUTOFF_DATE.date())
print("Elapsed calendar days in month:", elapsed_calendar_days)
print("Elapsed working days in month:", elapsed_working_days)
print("Remaining working days in month:", remaining_working_days)
print("Elapsed working days in year:", elapsed_year_working_days)
print("Remaining working days in year:", remaining_year_working_days)


# ==================================================================================================
# 2. STANDARD KEYS AND SAFE CALCULATION HELPERS
# ==================================================================================================

def brand_kpi_key(value):
    """Controlled brand key used to join target, shipment, stock, and depletion."""
    if pd.isna(value):
        return np.nan

    text = str(value).upper().strip()

    remove_words = [
        " BEER", " MALT", " CRATE", " 24X33", "24X33",
        " H+HK", "H+HK", " H + HK", "H + HK"
    ]

    for word in remove_words:
        text = text.replace(word, "")

    text = (
        text
        .replace("HABESHA KOSTARA TAEM", "HABESHA KOSTARA")
        .replace("KOSTARA TAEM", "HABESHA KOSTARA")
        .replace("KOSTARA", "HABESHA KOSTARA")
        .replace("  ", " ")
        .strip()
    )

    if "HABESHA KOSTARA" in text:
        return "HABESHA KOSTARA"
    if "HABESHA" in text:
        return "HABESHA"
    if "KIDAME" in text:
        return "KIDAME"
    if "FETA" in text:
        return "FETA"
    if "NEGUS" in text:
        return "NEGUS"

    return text


def brand_display_name(key):
    mapping = {
        "HABESHA": "Habesha",
        "FETA": "Feta",
        "NEGUS": "Negus",
        "KIDAME": "Kidame",
        "HABESHA KOSTARA": "Habesha Kostara",
    }

    if pd.isna(key):
        return np.nan

    return mapping.get(str(key).upper().strip(), str(key).title())


def safe_divide(numerator, denominator):
    return np.where(denominator == 0, 0, numerator / denominator)


def add_agent_division_fields(df, agent_col="Agent", dv_col=None):
    """
    Adds standardized Agent, Agent Key, and Division columns using Agent MDM.

    New logic:
    - First build a source key from the agent name found in the fact table.
    - Merge to Agent MDM using Agent Source Key.
    - Use Agent Reporting Name as the final Agent display name.
    - Rebuild Agent Key from Agent Reporting Name so S4H naming and depletion naming consolidate.
    """
    out = df.copy()

    if agent_col not in out.columns:
        out[agent_col] = np.nan

    out["Agent Source Name"] = out[agent_col]
    out["Agent Source Key"] = out["Agent Source Name"].apply(clean_text)

    if dv_col is not None and dv_col in out.columns:
        out["Division_Source"] = out[dv_col]
    else:
        out["Division_Source"] = np.nan

    out = out.merge(
        dim_agent_reporting,
        how="left",
        on="Agent Source Key"
    )

    out["Agent"] = out["Agent Reporting Name"].fillna(out["Agent Source Name"])
    out["Agent Key"] = out["Agent"].apply(clean_text)
    out["Division"] = out["DV_From_Agent_MDM"].fillna(out["Division_Source"])

    return out


def metric_only(df, metric_cols):
    """
    Returns one unique metric row per Agent Key + Brand KPI Key before merging.

    This prevents row multiplication in final_kpi_wide. Without this aggregation,
    duplicate keys in stock / target / PY / actual / depletion tables can multiply
    shipment values during chained merges and make the final total check wrong.
    """
    key_cols = ["Agent Key", "Brand KPI Key"]

    available_metric_cols = [
        col for col in metric_cols
        if col in df.columns
    ]

    work_cols = [
        col for col in key_cols + available_metric_cols
        if col in df.columns
    ]

    work = df[work_cols].copy()

    for key_col in key_cols:
        if key_col not in work.columns:
            work[key_col] = np.nan

    for col in available_metric_cols:
        work[col] = to_number(work[col]).fillna(0)

    work = (
        work
        .groupby(key_cols, dropna=False)[available_metric_cols]
        .sum()
        .reset_index()
    )

    return work


def ensure_numeric_columns(df, cols):
    for col in cols:
        if col not in df.columns:
            df[col] = 0
        df[col] = df[col].fillna(0)
    return df


def target_pivot(df, value_col, prefix):
    pivot = (
        df
        .pivot_table(
            index=["Division", "Agent Key", "Agent", "Brand KPI Key"],
            columns="Target Type",
            values=value_col,
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
        .rename(columns={
            "AOP": f"{prefix}_AOP",
            "SNOP": f"{prefix}_SNOP",
            "Sales Target": f"{prefix}_Sales"
        })
    )

    for col in [f"{prefix}_AOP", f"{prefix}_SNOP", f"{prefix}_Sales"]:
        if col not in pivot.columns:
            pivot[col] = 0

    return pivot


# ==================================================================================================
# 3. ACTUAL SHIPMENT: MTD AND YTD BY AGENT + BRAND
# ==================================================================================================

actual_ab = actual_to_date.copy()
actual_ab["Brand KPI Key"] = actual_ab["Reporting Name"].apply(brand_kpi_key)

actual_mtd_ab = (
    actual_ab[actual_ab["Shipment Date"].between(CURRENT_MONTH_START, LAST_SHIPMENT_DATE)]
    .groupby(["DV", "Agent Key", "Agent", "Brand KPI Key"], dropna=False)
    .agg(MTD_Shipment=("Actual Shipment Crates", "sum"))
    .reset_index()
    .rename(columns={"DV": "Division"})
)

actual_ytd_ab = (
    actual_ab[actual_ab["Shipment Date"].between(YEAR_START, LAST_SHIPMENT_DATE)]
    .groupby(["DV", "Agent Key", "Agent", "Brand KPI Key"], dropna=False)
    .agg(YTD_Shipment=("Actual Shipment Crates", "sum"))
    .reset_index()
    .rename(columns={"DV": "Division"})
)


# ==================================================================================================
# 4. PY SHIPMENT: MTD AND YTD BY AGENT + BRAND
# ==================================================================================================

py_ab = py_to_date.copy()
py_ab["Brand KPI Key"] = py_ab["Reporting Name"].apply(brand_kpi_key)

py_mtd_ab = (
    py_ab[py_ab["PY Shipment Date"].between(PY_MONTH_START, PY_CUTOFF_DATE)]
    .groupby(["DV", "Agent Key", "Agent", "Brand KPI Key"], dropna=False)
    .agg(MTD_PY_Shipment=("PY Shipment Crates", "sum"))
    .reset_index()
    .rename(columns={"DV": "Division"})
)

py_ytd_ab = (
    py_ab[py_ab["PY Shipment Date"].between(PY_YEAR_START, PY_CUTOFF_DATE)]
    .groupby(["DV", "Agent Key", "Agent", "Brand KPI Key"], dropna=False)
    .agg(YTD_PY_Shipment=("PY Shipment Crates", "sum"))
    .reset_index()
    .rename(columns={"DV": "Division"})
)


# ==================================================================================================
# 5. TARGETS: MTD, YTD, FULL MONTH, AND FULL YEAR BY AGENT + BRAND + TARGET TYPE
# ==================================================================================================

target_ab = fact_target.copy()
target_ab = add_agent_division_fields(target_ab, agent_col="Agent", dv_col="DV")
target_ab["Brand KPI Key"] = target_ab["Reporting Name"].apply(brand_kpi_key)

mtd_target_ab_long = (
    target_ab[target_ab["Month Start"] == CURRENT_MONTH_START]
    .groupby(["Division", "Agent Key", "Agent", "Brand KPI Key", "Target Type"], dropna=False)
    .agg(
        MTD_Target=("Target Crates To Date", "sum"),
        Full_Month_Target=("Target Crates", "sum")
    )
    .reset_index()
)

ytd_target_ab_long = (
    target_ab
    .groupby(["Division", "Agent Key", "Agent", "Brand KPI Key", "Target Type"], dropna=False)
    .agg(
        YTD_Target=("Target Crates To Date", "sum"),
        Full_Year_Target=("Target Crates", "sum")
    )
    .reset_index()
)

mtd_target_pivot = target_pivot(mtd_target_ab_long, "MTD_Target", "MTD Target")
full_month_target_pivot = target_pivot(mtd_target_ab_long, "Full_Month_Target", "Full Month Target")
ytd_target_pivot = target_pivot(ytd_target_ab_long, "YTD_Target", "YTD Target")
full_year_target_pivot = target_pivot(ytd_target_ab_long, "Full_Year_Target", "Full Year Target")


# ==================================================================================================
# 6. BEGINNING STOCK BY AGENT + BRAND
# ==================================================================================================

stock_work = fact_stock.copy()
stock_work = clean_columns(stock_work)

stock_date_col = find_col(
    stock_work,
    ["Reporting Date", "Date", "Stock Date", "Month", "Posting Date"],
    "fact_stock"
)

stock_work[stock_date_col] = pd.to_datetime(stock_work[stock_date_col], errors="coerce", dayfirst=True)
latest_stock_date = stock_work[stock_date_col].max()

if pd.isna(latest_stock_date):
    raise ValueError("Stock date is blank. Check stock Reporting Date column.")

stock_work = stock_work[stock_work[stock_date_col] == latest_stock_date].copy()

stock_id_cols = [col for col in ["Agent", stock_date_col, "Fact Source"] if col in stock_work.columns]
stock_value_cols = [col for col in stock_work.columns if col not in stock_id_cols]

if not stock_value_cols:
    raise ValueError("No stock brand columns found after excluding Agent, Reporting Date, and Fact Source.")

stock_long = stock_work.melt(
    id_vars=stock_id_cols,
    value_vars=stock_value_cols,
    var_name="Stock Brand Raw",
    value_name="Beginning_Stock"
)

stock_long["Beginning_Stock"] = to_number(stock_long["Beginning_Stock"]).fillna(0)
stock_long = stock_long[stock_long["Beginning_Stock"] != 0].copy()

stock_long["Stock Brand Clean"] = (
    stock_long["Stock Brand Raw"]
    .astype(str)
    .str.replace("Crate 24x33", "", regex=False)
    .str.replace("Crate 24X33", "", regex=False)
    .str.replace("CRATE 24X33", "", regex=False)
    .str.replace("crate 24x33", "", regex=False)
    .str.strip()
)

stock_long["Brand KPI Key"] = stock_long["Stock Brand Clean"].apply(brand_kpi_key)
stock_long = add_agent_division_fields(stock_long, agent_col="Agent", dv_col=None)

stock_ab = (
    stock_long
    .groupby(["Division", "Agent Key", "Agent", "Brand KPI Key"], dropna=False)
    .agg(Beginning_Stock=("Beginning_Stock", "sum"))
    .reset_index()
)


# ==================================================================================================
# 7. DEPLETION: MTD AND YTD BY AGENT + BRAND
# ==================================================================================================

depletion_work = fact_ytd_depletion.copy()
depletion_work = clean_columns(depletion_work)

depletion_date_col = find_col(
    depletion_work,
    ["Date", "Billing Date", "Depletion Date", "Posting Date", "Invoice Date"],
    "fact_ytd_depletion"
)

depletion_agent_col = find_col(
    depletion_work,
    ["Sales Unit", "Agent", "Sold-to Name", "Customer Name", "Customer"],
    "fact_ytd_depletion"
)

depletion_brand_col = find_col(
    depletion_work,
    ["Brand H+HK", "Product", "Item Description", "Brand", "Material Description"],
    "fact_ytd_depletion"
)

depletion_dv_col = None
for col in ["DV", "Division", "Region"]:
    if col in depletion_work.columns:
        depletion_dv_col = col
        break

depletion_hl_col = None
for col in ["Dep HL", "Depletion HL", "Depletion, HL", "Volume HL", "HL"]:
    if col in depletion_work.columns:
        depletion_hl_col = col
        break

depletion_crates_col = None
if depletion_hl_col is None:
    for col in ["Depletion", "Depletion Crates", "Invoiced Quantity", "Quantity", "Qty", "Crates", "Dep.", "Dep Crates"]:
        if col in depletion_work.columns:
            depletion_crates_col = col
            break

if depletion_crates_col is None and depletion_hl_col is None:
    raise KeyError(
        "No depletion volume column found. "
        f"Available depletion columns: {list(depletion_work.columns)}"
    )

depletion_work[depletion_date_col] = pd.to_datetime(depletion_work[depletion_date_col], errors="coerce", dayfirst=True)
depletion_work = depletion_work[depletion_work[depletion_date_col].notna()].copy()

depletion_work["Depletion Date"] = depletion_work[depletion_date_col]
depletion_work["Original Brand"] = depletion_work[depletion_brand_col]
depletion_work["Brand KPI Key"] = depletion_work["Original Brand"].apply(brand_kpi_key)

depletion_work = add_agent_division_fields(
    depletion_work,
    agent_col=depletion_agent_col,
    dv_col=depletion_dv_col
)

if depletion_hl_col is not None:
    depletion_work["Depletion_HL"] = to_number(depletion_work[depletion_hl_col])
    depletion_work["Depletion_Crates"] = depletion_work["Depletion_HL"] / HL_PER_CRATE
else:
    depletion_work["Depletion_Crates"] = to_number(depletion_work[depletion_crates_col])
    depletion_work["Depletion_HL"] = depletion_work["Depletion_Crates"] * HL_PER_CRATE

mtd_depletion_ab = (
    depletion_work[depletion_work["Depletion Date"].between(CURRENT_MONTH_START, LAST_SHIPMENT_DATE)]
    .groupby(["Division", "Agent Key", "Agent", "Brand KPI Key"], dropna=False)
    .agg(MTD_Depletion=("Depletion_Crates", "sum"))
    .reset_index()
)

ytd_depletion_ab = (
    depletion_work[depletion_work["Depletion Date"].between(YEAR_START, LAST_SHIPMENT_DATE)]
    .groupby(["Division", "Agent Key", "Agent", "Brand KPI Key"], dropna=False)
    .agg(YTD_Depletion=("Depletion_Crates", "sum"))
    .reset_index()
)


# ==================================================================================================
# 8. FINAL BASE FROM UNION OF ALL SOURCES
# ==================================================================================================

identity_sources = []
for df in [
    actual_mtd_ab,
    actual_ytd_ab,
    py_mtd_ab,
    py_ytd_ab,
    mtd_target_ab_long,
    ytd_target_ab_long,
    stock_ab,
    mtd_depletion_ab,
    ytd_depletion_ab,
]:
    keep_cols = ["Division", "Agent Key", "Agent", "Brand KPI Key"]
    identity_sources.append(df[[col for col in keep_cols if col in df.columns]].copy())

final_base_raw = pd.concat(identity_sources, ignore_index=True).drop_duplicates()

final_base = (
    final_base_raw
    .groupby(["Agent Key", "Brand KPI Key"], dropna=False)
    .agg(
        Division=("Division", lambda x: x.dropna().iloc[0] if len(x.dropna()) else np.nan),
        Agent=("Agent", lambda x: x.dropna().iloc[0] if len(x.dropna()) else np.nan)
    )
    .reset_index()
)

final_base["Brand"] = final_base["Brand KPI Key"].apply(brand_display_name)


# ==================================================================================================
# AGENT REPORTING NAME CONSOLIDATION CHECK
# ==================================================================================================

agent_reporting_consolidation_check = (
    final_base
    .groupby("Agent", dropna=False)
    .agg(
        Source_Key_Count=("Agent Key", "nunique"),
        Brand_Count=("Brand KPI Key", "nunique")
    )
    .reset_index()
    .sort_values(["Source_Key_Count", "Agent"], ascending=[False, True])
)

print("=" * 120)
print("AGENT REPORTING NAME CONSOLIDATION CHECK")
print("=" * 120)
print("Agent column now uses Agent Reporting Name from Agent MDM.")
display(agent_reporting_consolidation_check.head(20))


# ==================================================================================================
# 9. MERGE ALL METRICS INTO FINAL CRATES TABLE
# ==================================================================================================

final_kpi_wide = (
    final_base
    .merge(metric_only(stock_ab, ["Beginning_Stock"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(actual_mtd_ab, ["MTD_Shipment"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(py_mtd_ab, ["MTD_PY_Shipment"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(mtd_depletion_ab, ["MTD_Depletion"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(actual_ytd_ab, ["YTD_Shipment"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(py_ytd_ab, ["YTD_PY_Shipment"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(ytd_depletion_ab, ["YTD_Depletion"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(mtd_target_pivot, ["MTD Target_AOP", "MTD Target_SNOP", "MTD Target_Sales"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(ytd_target_pivot, ["YTD Target_AOP", "YTD Target_SNOP", "YTD Target_Sales"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(full_month_target_pivot, ["Full Month Target_AOP", "Full Month Target_SNOP", "Full Month Target_Sales"]), how="left", on=["Agent Key", "Brand KPI Key"])
    .merge(metric_only(full_year_target_pivot, ["Full Year Target_AOP", "Full Year Target_SNOP", "Full Year Target_Sales"]), how="left", on=["Agent Key", "Brand KPI Key"])
)

final_kpi_wide = final_kpi_wide.rename(columns={
    "MTD Target_AOP": "MTD Target AOP",
    "MTD Target_SNOP": "MTD Target SNOP",
    "MTD Target_Sales": "MTD Target Sales",
    "YTD Target_AOP": "YTD Target AOP",
    "YTD Target_SNOP": "YTD Target SNOP",
    "YTD Target_Sales": "YTD Target Sales",
    "Full Month Target_AOP": "Full Month Target AOP",
    "Full Month Target_SNOP": "Full Month Target SNOP",
    "Full Month Target_Sales": "Full Month Target Sales",
    "Full Year Target_AOP": "Full Year Target AOP",
    "Full Year Target_SNOP": "Full Year Target SNOP",
    "Full Year Target_Sales": "Full Year Target Sales",
})

volume_cols_crates = [
    "Beginning_Stock",
    "MTD_Shipment",
    "MTD_PY_Shipment",
    "MTD_Depletion",
    "YTD_Shipment",
    "YTD_PY_Shipment",
    "YTD_Depletion",
    "MTD Target AOP",
    "MTD Target SNOP",
    "MTD Target Sales",
    "YTD Target AOP",
    "YTD Target SNOP",
    "YTD Target Sales",
    "Full Month Target AOP",
    "Full Month Target SNOP",
    "Full Month Target Sales",
    "Full Year Target AOP",
    "Full Year Target SNOP",
    "Full Year Target Sales",
]

final_kpi_wide = ensure_numeric_columns(final_kpi_wide, volume_cols_crates)


# ==================================================================================================
# 10. CALCULATED KPIs
# ==================================================================================================

final_kpi_wide["Stock_Balance"] = (
    final_kpi_wide["Beginning_Stock"] +
    final_kpi_wide["MTD_Shipment"] -
    final_kpi_wide["MTD_Depletion"]
).clip(lower=0)

final_kpi_wide["Avg_Daily_Depletion"] = final_kpi_wide["MTD_Depletion"] / elapsed_calendar_days

final_kpi_wide["Days_of_Stock"] = safe_divide(
    final_kpi_wide["Stock_Balance"],
    final_kpi_wide["Avg_Daily_Depletion"]
)

final_kpi_wide["Avg_Daily_MTD_Shipment"] = final_kpi_wide["MTD_Shipment"] / elapsed_working_days

final_kpi_wide["MTD_Landing"] = (
    final_kpi_wide["MTD_Shipment"] +
    final_kpi_wide["Avg_Daily_MTD_Shipment"] * remaining_working_days
)

final_kpi_wide["Avg_Daily_YTD_Shipment"] = final_kpi_wide["YTD_Shipment"] / elapsed_year_working_days

final_kpi_wide["Yearly_Landing"] = (
    final_kpi_wide["YTD_Shipment"] +
    final_kpi_wide["Avg_Daily_YTD_Shipment"] * remaining_year_working_days
)

final_kpi_wide["MTD Variance vs PY"] = final_kpi_wide["MTD_Shipment"] - final_kpi_wide["MTD_PY_Shipment"]
final_kpi_wide["YTD Variance vs PY"] = final_kpi_wide["YTD_Shipment"] - final_kpi_wide["YTD_PY_Shipment"]

final_kpi_wide["MTD Achievement vs PY %"] = safe_divide(final_kpi_wide["MTD_Shipment"], final_kpi_wide["MTD_PY_Shipment"])
final_kpi_wide["YTD Achievement vs PY %"] = safe_divide(final_kpi_wide["YTD_Shipment"], final_kpi_wide["YTD_PY_Shipment"])

for target_name, target_col in [
    ("AOP", "MTD Target AOP"),
    ("SNOP", "MTD Target SNOP"),
    ("Sales", "MTD Target Sales"),
]:
    final_kpi_wide[f"MTD Achievement vs {target_name} %"] = safe_divide(final_kpi_wide["MTD_Shipment"], final_kpi_wide[target_col])

for target_name, target_col in [
    ("AOP", "YTD Target AOP"),
    ("SNOP", "YTD Target SNOP"),
    ("Sales", "YTD Target Sales"),
]:
    final_kpi_wide[f"YTD Achievement vs {target_name} %"] = safe_divide(final_kpi_wide["YTD_Shipment"], final_kpi_wide[target_col])

for target_name, target_col in [
    ("AOP", "Full Month Target AOP"),
    ("SNOP", "Full Month Target SNOP"),
    ("Sales", "Full Month Target Sales"),
]:
    final_kpi_wide[f"MTD Landing vs {target_name} %"] = safe_divide(final_kpi_wide["MTD_Landing"], final_kpi_wide[target_col])

for target_name, target_col in [
    ("AOP", "Full Year Target AOP"),
    ("SNOP", "Full Year Target SNOP"),
    ("Sales", "Full Year Target Sales"),
]:
    final_kpi_wide[f"Yearly Landing vs {target_name} %"] = safe_divide(final_kpi_wide["Yearly_Landing"], final_kpi_wide[target_col])


# ==================================================================================================
# 11. CLEAN FINAL CRATES TABLE
# ==================================================================================================

final_kpi_wide = final_kpi_wide.rename(columns={
    "Beginning_Stock": "Beginning Stock",
    "MTD_Shipment": "MTD Shipment",
    "MTD_PY_Shipment": "MTD PY Shipment",
    "MTD_Depletion": "MTD Depletion",
    "Stock_Balance": "Stock Balance",
    "Avg_Daily_Depletion": "Avg Daily Depletion",
    "Days_of_Stock": "Days of Stock",
    "MTD_Landing": "MTD Landing",
    "YTD_Shipment": "YTD Shipment",
    "YTD_PY_Shipment": "YTD PY Shipment",
    "YTD_Depletion": "YTD Depletion",
    "Yearly_Landing": "Yearly Landing",
    "Avg_Daily_MTD_Shipment": "Avg Daily MTD Shipment",
    "Avg_Daily_YTD_Shipment": "Avg Daily YTD Shipment",
})

final_columns = [
    "Division",
    "Agent",
    "Brand",
    "Beginning Stock",
    "MTD Shipment",
    "MTD PY Shipment",
    "MTD Achievement vs PY %",
    "MTD Variance vs PY",
    "MTD Depletion",
    "Beginning Stock",
    "Stock Balance",
    "Avg Daily Depletion",
    "Days of Stock",
    "MTD Landing",
    "MTD Target AOP",
    "MTD Achievement vs AOP %",
    "MTD Target SNOP",
    "MTD Achievement vs SNOP %",
    "MTD Target Sales",
    "MTD Achievement vs Sales %",
    "YTD Shipment",
    "YTD PY Shipment",
    "YTD Achievement vs PY %",
    "YTD Variance vs PY",
    "YTD Depletion",
    "YTD Target AOP",
    "YTD Achievement vs AOP %",
    "YTD Target SNOP",
    "YTD Achievement vs SNOP %",
    "YTD Target Sales",
    "YTD Achievement vs Sales %",
    "Yearly Landing",
    "Full Year Target AOP",
    "Yearly Landing vs AOP %",
    "Full Year Target SNOP",
    "Yearly Landing vs SNOP %",
    "Full Year Target Sales",
    "Yearly Landing vs Sales %",
    "Full Month Target AOP",
    "MTD Landing vs AOP %",
    "Full Month Target SNOP",
    "MTD Landing vs SNOP %",
    "Full Month Target Sales",
    "MTD Landing vs Sales %",
    "Avg Daily MTD Shipment",
    "Avg Daily YTD Shipment",
]

final_kpi_wide = final_kpi_wide[[col for col in final_columns if col in final_kpi_wide.columns]].copy()
final_kpi_wide = final_kpi_wide.sort_values(["Division", "Agent", "MTD Shipment"], ascending=[True, True, False])


# ==================================================================================================
# 11B. HARD RECONCILIATION CHECK BEFORE UOM EXPANSION
# final_kpi_wide is still in CRATES only here.
# Expected: Final YTD Shipment must equal actual_to_date total shipment.
# ==================================================================================================

source_ytd_shipment_crates = actual_to_date["Actual Shipment Crates"].sum()
final_ytd_shipment_crates = final_kpi_wide["YTD Shipment"].sum()

source_mtd_shipment_crates = actual_to_date[
    actual_to_date["Shipment Date"].between(CURRENT_MONTH_START, LAST_SHIPMENT_DATE)
]["Actual Shipment Crates"].sum()

final_mtd_shipment_crates = final_kpi_wide["MTD Shipment"].sum()

reconciliation_check = pd.DataFrame({
    "Metric": [
        "Source YTD Shipment Crates",
        "Final YTD Shipment Crates",
        "YTD Difference",
        "Source MTD Shipment Crates",
        "Final MTD Shipment Crates",
        "MTD Difference"
    ],
    "Value": [
        source_ytd_shipment_crates,
        final_ytd_shipment_crates,
        final_ytd_shipment_crates - source_ytd_shipment_crates,
        source_mtd_shipment_crates,
        final_mtd_shipment_crates,
        final_mtd_shipment_crates - source_mtd_shipment_crates
    ]
})

reconciliation_check_display = reconciliation_check.copy()
reconciliation_check_display["Value"] = (
    reconciliation_check_display["Value"]
    .round(0)
    .map("{:,.0f}".format)
)

print("=" * 120)
print("FINAL KPI RECONCILIATION CHECK - BEFORE UOM CONVERSION")
print("=" * 120)
display(reconciliation_check_display)

if round(final_ytd_shipment_crates, 0) != round(source_ytd_shipment_crates, 0):
    raise ValueError(
        f"YTD Shipment mismatch. "
        f"Source actual_to_date = {source_ytd_shipment_crates:,.0f}, "
        f"Final table = {final_ytd_shipment_crates:,.0f}. "
        f"This means a metric table still has duplicate Agent Key + Brand KPI Key rows before merging."
    )


# ==================================================================================================
# 12. CREATE FINAL REPORT WITH UOM SELECTOR
# ==================================================================================================

final_kpi_crates = final_kpi_wide.copy()
final_kpi_crates["UOM"] = "Crates"

final_kpi_hl = final_kpi_wide.copy()
final_kpi_hl["UOM"] = "HL"

volume_cols_for_uom = [
    "Beginning Stock",
    "MTD Shipment",
    "MTD PY Shipment",
    "MTD Variance vs PY",
    "MTD Depletion",
    "Beginning Stock",
    "Stock Balance",
    "Avg Daily Depletion",
    "MTD Landing",
    "MTD Target AOP",
    "MTD Target SNOP",
    "MTD Target Sales",
    "YTD Shipment",
    "YTD PY Shipment",
    "YTD Variance vs PY",
    "YTD Depletion",
    "YTD Target AOP",
    "YTD Target SNOP",
    "YTD Target Sales",
    "Yearly Landing",
    "Full Year Target AOP",
    "Full Year Target SNOP",
    "Full Year Target Sales",
    "Full Month Target AOP",
    "Full Month Target SNOP",
    "Full Month Target Sales",
    "Avg Daily MTD Shipment",
    "Avg Daily YTD Shipment",
]

for col in [col for col in volume_cols_for_uom if col in final_kpi_hl.columns]:
    final_kpi_hl[col] = final_kpi_hl[col] * HL_PER_CRATE

front_cols = ["Division", "Agent", "Brand", "UOM"]

final_kpi_report = pd.concat([final_kpi_crates, final_kpi_hl], ignore_index=True)
final_kpi_report = final_kpi_report[[col for col in front_cols if col in final_kpi_report.columns] + [col for col in final_kpi_report.columns if col not in front_cols]].copy()

print("=" * 120)
print("FINAL KPI REPORT WITH UOM SELECTOR")
print("=" * 120)
display(final_kpi_report.head(20))


# ==================================================================================================
# 13. DISPLAY-FORMATTED PREVIEW
# ==================================================================================================

final_kpi_preview = final_kpi_report.copy()

number_cols_display = [col for col in volume_cols_for_uom + ["Days of Stock"] if col in final_kpi_preview.columns]
percent_cols_display = [col for col in final_kpi_preview.columns if col.endswith("%")]

# Remove duplicate columns before display formatting.
# This prevents final_kpi_preview[col] from returning a DataFrame instead of a Series.
final_kpi_preview = final_kpi_preview.loc[:, ~final_kpi_preview.columns.duplicated()].copy()

for col in number_cols_display:
    if col in final_kpi_preview.columns:
        final_kpi_preview[col] = (
            pd.to_numeric(final_kpi_preview[col], errors="coerce")
            .fillna(0)
            .round(0)
            .map("{:,.0f}".format)
        )

for col in percent_cols_display:
    if col in final_kpi_preview.columns:
        final_kpi_preview[col] = (
            pd.to_numeric(final_kpi_preview[col], errors="coerce")
            .fillna(0)
            .map("{:.1%}".format)
        )

print("=" * 120)
print("FINAL KPI REPORT PREVIEW - FORMATTED")
print("=" * 120)
display(final_kpi_preview.head(50))


# ==================================================================================================
# 14. TOTAL CHECK BY UOM
# Do not total Crates and HL together. Always filter one UOM in visuals.
# ==================================================================================================

check_metrics = [
    "MTD Shipment",
    "MTD PY Shipment",
    "MTD Depletion",
    "YTD Shipment",
    "YTD PY Shipment",
    "YTD Depletion",
    "Stock Balance",
    "MTD Target AOP",
    "MTD Target SNOP",
    "MTD Target Sales",
    "YTD Target AOP",
    "YTD Target SNOP",
    "YTD Target Sales",
]

for selected_uom in ["Crates", "HL"]:
    check_df = final_kpi_report[final_kpi_report["UOM"] == selected_uom].copy()

    final_total_check = pd.DataFrame({
        "Metric": [metric for metric in check_metrics if metric in check_df.columns],
        "Value": [check_df[metric].sum() for metric in check_metrics if metric in check_df.columns],
        "UOM": selected_uom
    })

    final_total_check["Value"] = final_total_check["Value"].round(0).map("{:,.0f}".format)

    print("=" * 120)
    print(f"FINAL KPI TOTAL CHECK - {selected_uom}")
    print("=" * 120)
    display(final_total_check)


# ==================================================================================================
# 15. DATA QUALITY CHECKS FOR FINAL BASE
# ==================================================================================================

missing_division_check = final_kpi_report[
    final_kpi_report["Division"].isna() &
    (final_kpi_report["UOM"] == "Crates")
].copy()

print("=" * 120)
print("FINAL KPI ROWS WITH MISSING DIVISION - REVIEW AGENT MDM IF NOT BLANK")
print("=" * 120)
display(missing_division_check[["Division", "Agent", "Brand", "UOM"]].drop_duplicates())

# ============================================================================
# FINAL OUTPUT TABLE: final_kpi_report
#
# Final user-facing / dashboard extract:
#   final_kpi_report
#
# Export:
#   final_kpi_report.csv
#
# Grain:
#   Division + Agent + Brand + UOM + Target Type
#
# Important stock logic:
#   Stock_per_Agent = Stock Balance
#   Avg_Daily_Depletion = MTD Depletion / elapsed calendar days
#   Days of Stock = Stock_per_Agent / Avg_Daily_Depletion
#
# This section avoids reverse-calculating Avg Daily Depletion from Days Stock.
# ============================================================================

backend_kpi_report = final_kpi_report.copy()
backend_kpi_report = backend_kpi_report.loc[:, ~backend_kpi_report.columns.duplicated()].copy()

# --------------------------------------------------------------------------------------------------
# Column compatibility helpers
# --------------------------------------------------------------------------------------------------

def ensure_backend_col(df, target_col, candidate_cols, default_value=0):
    """
    Ensures target_col exists by copying the first available candidate column.
    This keeps the final extract stable even if earlier sections use either
    space-based names or underscore-based names.
    """
    if target_col in df.columns:
        return df

    for col in candidate_cols:
        if col in df.columns:
            df[target_col] = df[col]
            return df

    df[target_col] = default_value
    return df


backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "Beginning Stock",
    ["Beginning Stock", "Beginning_Stock"],
    0
)

backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "Stock Balance",
    ["Stock Balance", "Stock_Balance"],
    0
)

backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "Avg Daily Depletion",
    ["Avg Daily Depletion", "Avg_Daily_Depletion"],
    0
)

backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "MTD Shipment",
    ["MTD Shipment", "MTD_Shipment"],
    0
)

backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "YTD Shipment",
    ["YTD Shipment", "YTD_Shipment"],
    0
)

backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "MTD Depletion",
    ["MTD Depletion", "MTD_Depletion"],
    0
)

backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "YTD Depletion",
    ["YTD Depletion", "YTD_Depletion"],
    0
)

backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "MTD PY Shipment",
    ["MTD PY Shipment", "MTD_PY_Shipment"],
    0
)

backend_kpi_report = ensure_backend_col(
    backend_kpi_report,
    "YTD PY Shipment",
    ["YTD PY Shipment", "YTD_PY_Shipment"],
    0
)

required_final_base_cols = [
    "Division",
    "Agent",
    "Brand",
    "UOM",
    "MTD Shipment",
    "YTD Shipment",
    "MTD Depletion",
    "YTD Depletion",
    "MTD PY Shipment",
    "YTD PY Shipment",
    "Beginning Stock",
    "Stock Balance",
    "Avg Daily Depletion",
]

for col in required_final_base_cols:
    if col not in backend_kpi_report.columns:
        backend_kpi_report[col] = 0

# Ensure numeric source fields are numeric before aggregation.
numeric_backend_cols = [
    "MTD Shipment",
    "YTD Shipment",
    "MTD Depletion",
    "YTD Depletion",
    "MTD PY Shipment",
    "YTD PY Shipment",
    "Beginning Stock",
    "Stock Balance",
    "Avg Daily Depletion",
    "MTD Target Sales",
    "YTD Target Sales",
    "MTD Target SNOP",
    "YTD Target SNOP",
    "MTD Target AOP",
    "YTD Target AOP",
]

for col in numeric_backend_cols:
    if col in backend_kpi_report.columns:
        backend_kpi_report[col] = pd.to_numeric(
            backend_kpi_report[col],
            errors="coerce"
        ).fillna(0)

target_type_map = {
    "Sales": {
        "MTD Target": "MTD Target Sales",
        "YTD Target": "YTD Target Sales",
    },
    "SNOP": {
        "MTD Target": "MTD Target SNOP",
        "YTD Target": "YTD Target SNOP",
    },
    "AOP": {
        "MTD Target": "MTD Target AOP",
        "YTD Target": "YTD Target AOP",
    },
}

final_report_tables = []

for uom_value in sorted(backend_kpi_report["UOM"].dropna().unique()):

    uom_source = backend_kpi_report[
        backend_kpi_report["UOM"] == uom_value
    ].copy()

    for target_type, target_cols in target_type_map.items():

        mtd_target_col = target_cols["MTD Target"]
        ytd_target_col = target_cols["YTD Target"]

        if mtd_target_col not in uom_source.columns:
            uom_source[mtd_target_col] = 0

        if ytd_target_col not in uom_source.columns:
            uom_source[ytd_target_col] = 0

        temp_table = (
            uom_source
            .groupby(["Division", "Agent", "Brand", "UOM"], dropna=False)
            .agg(
                MTD_Sales=("MTD Shipment", "sum"),
                MTD_Target=(mtd_target_col, "sum"),
                MTD_Depletion=("MTD Depletion", "sum"),
                YTD_Sales=("YTD Shipment", "sum"),
                YTD_Target=(ytd_target_col, "sum"),
                YTD_Depletion=("YTD Depletion", "sum"),
                MTD_PY=("MTD PY Shipment", "sum"),
                YTD_PY=("YTD PY Shipment", "sum"),
                Beginning_Stock=("Beginning Stock", "sum"),
                Stock_per_Agent=("Stock Balance", "sum"),
                Avg_Daily_Depletion=("Avg Daily Depletion", "sum")
            )
            .reset_index()
        )

        temp_table["Target Type"] = target_type

        temp_table["MTD Ach"] = safe_divide(
            temp_table["MTD_Sales"],
            temp_table["MTD_Target"]
        ) - 1

        temp_table["YTD Ach"] = safe_divide(
            temp_table["YTD_Sales"],
            temp_table["YTD_Target"]
        ) - 1

        temp_table["MTD Sales vs PY"] = safe_divide(
            temp_table["MTD_Sales"],
            temp_table["MTD_PY"]
        ) - 1

        temp_table["YTD VS PY Ach"] = safe_divide(
            temp_table["YTD_Sales"],
            temp_table["YTD_PY"]
        ) - 1

        temp_table["MTD Depletion Ach"] = safe_divide(
            temp_table["MTD_Depletion"],
            temp_table["MTD_Target"]
        ) - 1

        temp_table["YTD Depletion Ach"] = safe_divide(
            temp_table["YTD_Depletion"],
            temp_table["YTD_Target"]
        ) - 1

        # Correct stock coverage.
        temp_table["Days of Stock"] = safe_divide(
            temp_table["Stock_per_Agent"],
            temp_table["Avg_Daily_Depletion"]
        )

        temp_table["Stock Level Flag"] = np.select(
            [
                temp_table["Days of Stock"] > 10,
                temp_table["Days of Stock"].between(6, 10, inclusive="both"),
                temp_table["Days of Stock"] < 6
            ],
            [
                "Over-stocked",
                "Optimal",
                "Low Stock"
            ],
            default="No Depletion"
        )

        final_report_tables.append(temp_table)

final_kpi_report = pd.concat(
    final_report_tables,
    ignore_index=True
)

# --------------------------------------------------------------------------------------------------
# Text cleanup to avoid duplicated filter values like "West" vs "WEST" or "North-East" vs "North East"
# --------------------------------------------------------------------------------------------------

def proper_text(value):
    if pd.isna(value):
        return np.nan

    text = (
        str(value)
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
        .replace("-", " ")
        .strip()
    )

    while "  " in text:
        text = text.replace("  ", " ")

    if text == "" or text.upper() in ["NAN", "NONE"]:
        return np.nan

    return text.title()


def normalize_uom(value):
    text = str(value).strip().upper()

    if text == "HL":
        return "HL"

    if text == "CRATES":
        return "Crates"

    return str(value).strip()


def normalize_target_type(value):
    text = str(value).strip().upper()

    if text in ["SALES", "SALES TARGET"]:
        return "Sales"

    if text == "SNOP":
        return "SNOP"

    if text == "AOP":
        return "AOP"

    return str(value).strip()


final_kpi_report["Division"] = final_kpi_report["Division"].apply(proper_text)
final_kpi_report["Agent"] = final_kpi_report["Agent"].apply(proper_text)
final_kpi_report["Brand"] = final_kpi_report["Brand"].apply(proper_text)
final_kpi_report["UOM"] = final_kpi_report["UOM"].apply(normalize_uom)
final_kpi_report["Target Type"] = final_kpi_report["Target Type"].apply(normalize_target_type)

# Re-aggregate after text cleanup, in case rows became identical after normalization.
final_kpi_report = (
    final_kpi_report
    .groupby(
        ["Division", "Agent", "Brand", "UOM", "Target Type"],
        dropna=False
    )
    .agg(
        MTD_Sales=("MTD_Sales", "sum"),
        MTD_Target=("MTD_Target", "sum"),
        MTD_Depletion=("MTD_Depletion", "sum"),
        YTD_Sales=("YTD_Sales", "sum"),
        YTD_Target=("YTD_Target", "sum"),
        YTD_Depletion=("YTD_Depletion", "sum"),
        MTD_PY=("MTD_PY", "sum"),
        YTD_PY=("YTD_PY", "sum"),
        Beginning_Stock=("Beginning_Stock", "sum"),
        Stock_per_Agent=("Stock_per_Agent", "sum"),
        Avg_Daily_Depletion=("Avg_Daily_Depletion", "sum")
    )
    .reset_index()
)

# Recalculate all ratios after re-aggregation.
final_kpi_report["MTD Ach"] = safe_divide(
    final_kpi_report["MTD_Sales"],
    final_kpi_report["MTD_Target"]
) - 1

final_kpi_report["YTD Ach"] = safe_divide(
    final_kpi_report["YTD_Sales"],
    final_kpi_report["YTD_Target"]
) - 1

final_kpi_report["MTD Sales vs PY"] = safe_divide(
    final_kpi_report["MTD_Sales"],
    final_kpi_report["MTD_PY"]
) - 1

final_kpi_report["YTD VS PY Ach"] = safe_divide(
    final_kpi_report["YTD_Sales"],
    final_kpi_report["YTD_PY"]
) - 1

final_kpi_report["MTD Depletion Ach"] = safe_divide(
    final_kpi_report["MTD_Depletion"],
    final_kpi_report["MTD_Target"]
) - 1

final_kpi_report["YTD Depletion Ach"] = safe_divide(
    final_kpi_report["YTD_Depletion"],
    final_kpi_report["YTD_Target"]
) - 1

final_kpi_report["Days of Stock"] = safe_divide(
    final_kpi_report["Stock_per_Agent"],
    final_kpi_report["Avg_Daily_Depletion"]
)

final_kpi_report["Stock Level Flag"] = np.select(
    [
        final_kpi_report["Days of Stock"] > 10,
        final_kpi_report["Days of Stock"].between(6, 10, inclusive="both"),
        final_kpi_report["Days of Stock"] < 6
    ],
    [
        "Over-stocked",
        "Optimal",
        "Low Stock"
    ],
    default="No Depletion"
)

final_kpi_report = final_kpi_report[
    [
        "Division",
        "Agent",
        "Brand",
        "UOM",
        "Target Type",
        "MTD_Sales",
        "MTD_Target",
        "MTD Ach",
        "MTD_Depletion",
        "MTD Depletion Ach",
        "YTD_Sales",
        "YTD_Target",
        "YTD Ach",
        "YTD_Depletion",
        "YTD Depletion Ach",
        "MTD_PY",
        "MTD Sales vs PY",
        "YTD_PY",
        "YTD VS PY Ach",
        "Beginning_Stock",
        "Stock_per_Agent",
        "Avg_Daily_Depletion",
        "Days of Stock",
        "Stock Level Flag"
    ]
].copy()

final_kpi_report = final_kpi_report.sort_values(
    by=["UOM", "Target Type", "Division", "Agent", "Brand", "MTD_Sales"],
    ascending=[True, True, True, True, True, False]
).reset_index(drop=True)

# --------------------------------------------------------------------------------------------------
# Display preview only. Export stays numeric.
# --------------------------------------------------------------------------------------------------

final_kpi_report_display = final_kpi_report.copy()
final_kpi_report_display = final_kpi_report_display.loc[
    :,
    ~final_kpi_report_display.columns.duplicated()
].copy()

number_cols_final = [
    "MTD_Sales",
    "MTD_Target",
    "MTD_Depletion",
    "YTD_Sales",
    "YTD_Target",
    "YTD_Depletion",
    "MTD_PY",
    "YTD_PY",
    "Beginning_Stock",
    "Stock_per_Agent",
    "Avg_Daily_Depletion",
    "Days of Stock"
]

growth_cols_final = [
    "MTD Ach",
    "MTD Depletion Ach",
    "YTD Ach",
    "YTD Depletion Ach",
    "MTD Sales vs PY",
    "YTD VS PY Ach"
]

for col in number_cols_final:
    if col in final_kpi_report_display.columns:
        final_kpi_report_display[col] = (
            pd.to_numeric(final_kpi_report_display[col], errors="coerce")
            .fillna(0)
            .round(2)
            .map("{:,.2f}".format)
        )

for col in growth_cols_final:
    if col in final_kpi_report_display.columns:
        final_kpi_report_display[col] = (
            pd.to_numeric(final_kpi_report_display[col], errors="coerce")
            .fillna(0)
            .map("{:+.1%}".format)
        )

print("=" * 120)
print("FINAL DASHBOARD / REPORTING EXTRACT: final_kpi_report")
print("=" * 120)
print("This is the final extract to use as dashboard data.")
print("Included UOMs:", sorted(final_kpi_report["UOM"].dropna().unique()))
print("Included Target Types:", sorted(final_kpi_report["Target Type"].dropna().unique()))
print("Days of Stock formula: Stock_per_Agent / Avg_Daily_Depletion")
display(final_kpi_report_display.head(100))

# --------------------------------------------------------------------------------------------------
# Final table total check
# --------------------------------------------------------------------------------------------------

final_total_check = (
    final_kpi_report
    .groupby(["UOM", "Target Type"], dropna=False)
    .agg(
        MTD_Sales=("MTD_Sales", "sum"),
        MTD_Target=("MTD_Target", "sum"),
        MTD_Depletion=("MTD_Depletion", "sum"),
        YTD_Sales=("YTD_Sales", "sum"),
        YTD_Target=("YTD_Target", "sum"),
        YTD_Depletion=("YTD_Depletion", "sum"),
        MTD_PY=("MTD_PY", "sum"),
        YTD_PY=("YTD_PY", "sum"),
        Beginning_Stock=("Beginning_Stock", "sum"),
        Stock_per_Agent=("Stock_per_Agent", "sum"),
        Avg_Daily_Depletion=("Avg_Daily_Depletion", "sum")
    )
    .reset_index()
)

final_total_check["MTD Ach"] = safe_divide(
    final_total_check["MTD_Sales"],
    final_total_check["MTD_Target"]
) - 1

final_total_check["YTD Ach"] = safe_divide(
    final_total_check["YTD_Sales"],
    final_total_check["YTD_Target"]
) - 1

final_total_check["Days of Stock"] = safe_divide(
    final_total_check["Stock_per_Agent"],
    final_total_check["Avg_Daily_Depletion"]
)

final_total_check_display = final_total_check.copy()
final_total_check_display = final_total_check_display.loc[
    :,
    ~final_total_check_display.columns.duplicated()
].copy()

for col in [
    "MTD_Sales",
    "MTD_Target",
    "MTD_Depletion",
    "YTD_Sales",
    "YTD_Target",
    "YTD_Depletion",
    "MTD_PY",
    "YTD_PY",
    "Beginning_Stock",
    "Stock_per_Agent",
    "Avg_Daily_Depletion",
    "Days of Stock"
]:
    if col in final_total_check_display.columns:
        final_total_check_display[col] = (
            pd.to_numeric(final_total_check_display[col], errors="coerce")
            .fillna(0)
            .round(2)
            .map("{:,.2f}".format)
        )

for col in ["MTD Ach", "YTD Ach"]:
    if col in final_total_check_display.columns:
        final_total_check_display[col] = (
            pd.to_numeric(final_total_check_display[col], errors="coerce")
            .fillna(0)
            .map("{:+.1%}".format)
        )

print("=" * 120)
print("FINAL EXTRACT TOTAL CHECK BY UOM AND TARGET TYPE")
print("=" * 120)
display(final_total_check_display)

# --------------------------------------------------------------------------------------------------
# Export only the final extract.
# --------------------------------------------------------------------------------------------------

final_output_file = BASE_FOLDER / "final_kpi_report.csv"
final_kpi_report = final_kpi_report.loc[:, ~final_kpi_report.columns.duplicated()].copy()
final_kpi_report.to_csv(final_output_file, index=False, encoding="utf-8-sig")

print("=" * 120)
print("FINAL EXTRACT EXPORTED")
print("=" * 120)
print("Exported file:")
print(final_output_file)
print("Use final_kpi_report.csv as the data source for the HTML dashboard.")
print("Reminder: Filter UOM and Target Type before checking totals.")
print("Reminder: Days of Stock = Stock_per_Agent / Avg_Daily_Depletion.")

# %%
