import pandas as pd
import numpy as np
import re
import os
from datetime import datetime

OUTPUT_XLSX = "Personal_Expense_Tracker.xlsx"

# =========================
# Helpers
# =========================
def ensure_xlsx(file_path: str) -> str:
    """Convert .xls to .xlsx if needed."""
    base, ext = os.path.splitext(file_path)
    if ext.lower() == ".xls":
        new_path = base + ".xlsx"
        df = pd.read_excel(file_path, dtype=str)
        df.to_excel(new_path, index=False)
        print(f"Converted {file_path} -> {new_path}")
        return new_path
    return file_path

def to_num(x):
    if pd.isna(x): return np.nan
    s = str(x).strip().replace(",", "")
    if s == "": return np.nan
    if re.fullmatch(r"\(.*\)", s):
        try: return -float(s.strip("()"))
        except: return np.nan
    s = s.replace("Dr", "").replace("CR", "").replace("Cr", "").strip()
    try: return float(s)
    except: return np.nan

def parse_date(x):
    """Robust date parser (handles dd,mm,yyyy)."""
    if pd.isna(x): return pd.NaT
    if isinstance(x, (datetime, np.datetime64)): return pd.to_datetime(x)

    s = str(x).strip()
    s_norm = s.replace(",", "-").replace(".", "-").replace("/", "-")
    fmts = ["%d,%m,%Y","%d,%m,%y","%d-%m-%Y","%d-%m-%y",
            "%d/%m/%Y","%d/%m/%y","%d.%m.%Y","%d.%m.%y",
            "%d %b %Y","%d-%b-%Y","%Y-%m-%d"]
    for fmt in fmts:
        try:
            src = s if "%," in fmt else s_norm
            return datetime.strptime(src, fmt)
        except: pass
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def detect_bank(file_name: str) -> str:
    name = file_name.lower()
    if "optransactionhistory" in name or "icici" in name: return "ICICI"
    if "918010053388907" in name or "axis" in name: return "AXIS"
    if "3895" in name: return "HDFC1"
    if "7671" in name: return "HDFC2"
    return "UNKNOWN"

# =========================
# Header utilities
# =========================
def find_header_row(path, keywords, search_rows=60):
    raw = pd.read_excel(path, header=None, nrows=search_rows)
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).str.lower().tolist()
        if all(any(k in c for c in row) for k in keywords):
            return i
    return None

def pick_column(df, options):
    cols = [str(c).strip() for c in df.columns]
    lower = [c.lower() for c in cols]
    for opt in options:
        for i, lc in enumerate(lower):
            if opt in lc:
                return cols[i]
    return None

# =========================
# Parsers
# =========================
def parse_axis(path):
    hdr = find_header_row(path, ["tran date","particulars","dr","cr","bal"], 80) or 0
    df = pd.read_excel(path, header=hdr)
    return pd.DataFrame({
        "Date": df.get(pick_column(df, ["tran date","date"])),
        "Description": df.get(pick_column(df, ["particulars","description"])),
        "Ref": df.get(pick_column(df, ["chq","ref"])),
        "Debit": df.get(pick_column(df, ["dr","debit","withdrawal"])),
        "Credit": df.get(pick_column(df, ["cr","credit","deposit"])),
        "Balance": df.get(pick_column(df, ["bal","balance"]))
    })

def parse_hdfc(path):
    hdr = find_header_row(path, ["date","narration","withdrawal","deposit","balance"], 80) or 20
    df = pd.read_excel(path, header=hdr)
    return pd.DataFrame({
        "Date": df.get(pick_column(df, ["date"])),
        "Description": df.get(pick_column(df, ["narration","description"])),
        "Ref": df.get(pick_column(df, ["ref","cheque"])),
        "Debit": df.get(pick_column(df, ["withdrawal","debit"])),
        "Credit": df.get(pick_column(df, ["deposit","credit"])),
        "Balance": df.get(pick_column(df, ["balance"]))
    })

def parse_icici(path):
    hdr = find_header_row(path, ["transaction date","transaction remarks","withdrawal","deposit","balance"], 80) or 12
    df = pd.read_excel(path, header=hdr)
    return pd.DataFrame({
        "Date": df.get(pick_column(df, ["transaction date","date"])),
        "Description": df.get(pick_column(df, ["transaction remarks","description","particulars"])),
        "Ref": df.get(pick_column(df, ["cheque","ref"])),
        "Debit": df.get(pick_column(df, ["withdrawal","debit"])),
        "Credit": df.get(pick_column(df, ["deposit","credit"])),
        "Balance": df.get(pick_column(df, ["balance"]))
    })

def parse_fallback(path, bank="UNKNOWN"):
    df = pd.read_excel(path)
    return pd.DataFrame({
        "Date": df.get(pick_column(df, ["date"])),
        "Description": df.get(pick_column(df, ["description","narration","remarks"])),
        "Ref": df.get(pick_column(df, ["ref","chq"])),
        "Debit": df.get(pick_column(df, ["debit","withdrawal"])),
        "Credit": df.get(pick_column(df, ["credit","deposit"])),
        "Balance": df.get(pick_column(df, ["balance"]))
    })

# =========================
# Wrapper
# =========================
def parse_file(path, logs):
    path = ensure_xlsx(path)
    bank = detect_bank(os.path.basename(path))
    print(f"Parsing {os.path.basename(path)} as {bank}...")
    try:
        if bank=="AXIS": df = parse_axis(path)
        elif bank in ("HDFC1","HDFC2"): df = parse_hdfc(path)
        elif bank=="ICICI": df = parse_icici(path)
        else: df = parse_fallback(path, bank)
    except Exception as e:
        print(f"Schema parser failed for {bank}: {e}")
        df = parse_fallback(path, bank)

    # Clean
    df["Date"] = df["Date"].apply(parse_date)
    df["Debit"] = df["Debit"].apply(to_num)
    df["Credit"] = df["Credit"].apply(to_num)
    df["Balance"] = df["Balance"].apply(to_num)
    df["Bank"] = bank

    before=len(df)
    df=df.dropna(subset=["Date"])
    after=len(df)
    logs.append({"File": os.path.basename(path),"Bank":bank,"RowsParsed":after,"RowsDropped":before-after})
    return df

# =========================
# Dashboard
# =========================
CATEGORY_LIST = ["Groceries","Dining","Shopping","Utilities","Rent/Home",
                 "Transport","Fees/Charges","Salary","Transfers",
                 "Investments","Insurance","Interest","Income-Other","Expense-Other"]

def build_dashboard(consolidated, logs, output_path):
    consolidated = consolidated.sort_values("Date").reset_index(drop=True)

    consolidated["Type"] = np.where(consolidated["Debit"].fillna(0)>0,"Expense",
                                    np.where(consolidated["Credit"].fillna(0)>0,"Income","Other"))
    consolidated["Amount"] = consolidated["Credit"].fillna(0)-consolidated["Debit"].fillna(0)
    consolidated["Month"] = consolidated["Date"].dt.strftime("%b-%Y")

    # Insert Category column after Description
    cols = list(consolidated.columns)
    insert_at = cols.index("Description")+1
    consolidated.insert(insert_at,"Category","")

    summary_cat = consolidated.pivot_table(index="Category",columns="Type",values="Amount",aggfunc="sum",fill_value=0).reset_index()
    summary_month = (consolidated.groupby(["Month","Type"],as_index=False)["Amount"]
                     .sum().pivot(index="Month",columns="Type",values="Amount").fillna(0).reset_index())
    log_df = pd.DataFrame(logs)

    with pd.ExcelWriter(output_path,engine="xlsxwriter",datetime_format="dd-mmm-yyyy") as writer:
        consolidated.to_excel(writer,"Transactions",index=False)
        summary_cat.to_excel(writer,"Summary_By_Category",index=False)
        summary_month.to_excel(writer,"Summary_By_Month",index=False)
        pd.DataFrame([{"Category":c} for c in CATEGORY_LIST]).to_excel(writer,"Category_List",index=False)
        log_df.to_excel(writer,"Parse_Log",index=False)

        wb=writer.book
        ws_t=writer.sheets["Transactions"]

        # Dropdown validation for Category column
        dv_range=f"Category_List!$A$2:$A${len(CATEGORY_LIST)+1}"
        last_row=len(consolidated)+1
        ws_t.data_validation(1,insert_at,last_row,insert_at,{
            "validate":"list",
            "source":dv_range,
            "input_message":"Pick category"
        })

# =========================
# Main
# =========================
def main():
    files=[f for f in os.listdir(".") if f.lower().endswith((".xls",".xlsx")) and f!=OUTPUT_XLSX and not f.startswith("~$")]
    frames,logs=[],[]
    for f in files:
        try: frames.append(parse_file(f,logs))
        except Exception as e: print(f"Skipping {f}: {e}")
    if not frames: return
    consolidated=pd.concat(frames,ignore_index=True)
    build_dashboard(consolidated,logs,OUTPUT_XLSX)
    print(f"Saved -> {OUTPUT_XLSX}")

if __name__=="__main__":
    main()
