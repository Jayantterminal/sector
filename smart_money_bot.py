import os
import glob
import datetime
import requests
import warnings
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

BASE_DIR = r"D:\Bhavdata\bhavdata_3months_arranged"
DASHBOARD_FILE = r"D:\Bhavdata\SmartMoney_Live_Dashboard.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "*/*"
}

# =====================================================================
# 1. READ PREVIOUS USER INPUTS (ENTERED, BUY PRICE, DROP)
# =====================================================================
user_trade_actions = {}
user_buy_prices = {}
prev_tracking_dict = {}

if os.path.exists(DASHBOARD_FILE):
    try:
        prev_df = pd.read_excel(DASHBOARD_FILE, sheet_name="Top_Accumulation_Radar", skiprows=2)
        cols = [str(c).strip() for c in prev_df.columns]
        prev_df.columns = cols
        
        if "Symbol" in prev_df.columns:
            for _, r in prev_df.iterrows():
                sym = str(r["Symbol"]).strip().upper()
                if "First Detected Date" in prev_df.columns and pd.notna(r["First Detected Date"]):
                    prev_tracking_dict[sym] = str(r["First Detected Date"])
                if "Trade Action" in prev_df.columns and pd.notna(r["Trade Action"]):
                    act = str(r["Trade Action"]).strip().upper()
                    if act and act != "NAN":
                        user_trade_actions[sym] = act
                if "My Buy Price (Rs)" in prev_df.columns and pd.notna(r["My Buy Price (Rs)"]):
                    try:
                        bp = float(r["My Buy Price (Rs)"])
                        if bp > 0:
                            user_buy_prices[sym] = bp
                    except Exception:
                        pass
    except Exception as e:
        print(f"⚠️ Notice reading previous inputs: {e}")

# =====================================================================
# 2. HINGLISH NEWS & CATALYST FETCHER
# =====================================================================
def get_hinglish_catalyst(symbol, sector):
    try:
        url = f"https://news.google.com/rss/search?q={symbol}+share+OR+stock+market&hl=en-IN&gl=IN&ceid=IN:en"
        res = requests.get(url, headers=HEADERS, timeout=3)
        if res.status_code == 200:
            root = ET.fromstring(res.content)
            items = root.findall('.//item')
            if items:
                title = items[0].find('title').text or ""
                t_lower = title.lower()
                if any(w in t_lower for w in ["order", "bagged", "contract", "nhai", "project"]):
                    return "Naya major order / project deal announce hui hai"
                elif any(w in t_lower for w in ["block deal", "stake", "bulk deal", "promoter", "fii"]):
                    return "Institutional block deal / heavy stake accumulation"
                elif any(w in t_lower for w in ["profit", "result", "q1", "q2", "q3", "revenue", "quarter"]):
                    return "Quarterly financial results / business update news"
                elif any(w in t_lower for w in ["target", "upgrade", "buy", "brokerage"]):
                    return "Brokerage houses se target upgrade & buy call"
                elif any(w in t_lower for w in ["dividend", "split", "bonus"]):
                    return "Corporate action (Dividend / Split / Bonus) trigger"
                else:
                    return f"Headline: {title[:55]}..."
    except Exception:
        pass
    return f"{sector} me strong institutional delivery build-up"

# =====================================================================
# 3. TODAY'S BHAVCOPY AUTO-DOWNLOAD
# =====================================================================
today = datetime.date.today()
today_str = today.strftime("%d%m%Y")
month_folder_name = today.strftime("%Y-%m (%b)")
target_folder = os.path.join(BASE_DIR, month_folder_name)
os.makedirs(target_folder, exist_ok=True)

bhav_url = f"https://nsearchives.nseindia.com/products/content/sec_bhavdata_full_{today_str}.csv"
target_file_path = os.path.join(target_folder, f"sec_bhavdata_full_{today_str}.csv")

if not os.path.exists(target_file_path):
    print(f"⏳ Downloading Bhavcopy for today ({today_str})...")
    try:
        r = requests.get(bhav_url, headers=HEADERS, timeout=15)
        if r.status_code == 200 and len(r.content) > 5000:
            with open(target_file_path, "wb") as f:
                f.write(r.content)
            print(f"✅ Downloaded & saved to {month_folder_name}")
        else:
            print("ℹ️ Today's Bhavcopy not yet on NSE. Processing latest available files...")
    except Exception as e:
        print(f"⚠️ Download notice: {e}")
else:
    print(f"✅ Today's Bhavcopy CSV is up-to-date in {month_folder_name}")

# =====================================================================
# 4. NSE SECTOR MASTER & UNIVERSE
# =====================================================================
sector_mapping = {}
target_universe = set()

try:
    u500 = "https://nsearchives.nseindia.com/content/indices/ind_nifty500list.csv"
    r500 = requests.get(u500, headers=HEADERS, timeout=10)
    df_500 = pd.read_csv(pd.io.common.BytesIO(r500.content))
    for _, row in df_500.iterrows():
        sym = str(row["Symbol"]).strip().upper()
        sector_mapping[sym] = str(row.get("Industry", "Diversified")).strip()
        target_universe.add(sym)
except Exception:
    pass

try:
    umicro = "https://nsearchives.nseindia.com/content/indices/ind_niftymicrocap250_list.csv"
    rmicro = requests.get(umicro, headers=HEADERS, timeout=10)
    df_micro = pd.read_csv(pd.io.common.BytesIO(rmicro.content))
    for _, row in df_micro.iterrows():
        sym = str(row["Symbol"]).strip().upper()
        if sym not in sector_mapping:
            sector_mapping[sym] = str(row.get("Industry", "Microcap / Others")).strip()
        target_universe.add(sym)
except Exception:
    pass

if "CUPID" not in sector_mapping:
    sector_mapping["CUPID"] = "Healthcare / Consumer"
    target_universe.add("CUPID")

# =====================================================================
# 5. BHAVCOPY DATA MERGE & CLEAN
# =====================================================================
all_files = sorted(glob.glob(f"{BASE_DIR}/**/*.csv", recursive=True))
if not all_files:
    raise SystemExit("No CSV files found in directory!")

df_list = []
for f in all_files:
    try:
        tmp = pd.read_csv(f)
        tmp.columns = [c.strip().upper() for c in tmp.columns]
        if "SERIES" in tmp.columns:
            tmp = tmp[tmp["SERIES"].str.strip() == "EQ"]
        df_list.append(tmp)
    except Exception:
        continue

df = pd.concat(df_list, ignore_index=True)
if len(target_universe) > 10:
    df = df[df["SYMBOL"].str.strip().str.upper().isin(target_universe)].copy()

etf_keywords = ["BEES", "ETF", "10BE", "GOLD", "LIQUID", "NIFTY", "SENSEX", "MON100"]
df = df[~df["SYMBOL"].str.upper().str.contains("|".join(etf_keywords))].copy()

df["DATE1"] = pd.to_datetime(df["DATE1"].str.strip(), format="%d-%b-%Y")
df = df.sort_values(["SYMBOL", "DATE1"]).reset_index(drop=True)

for col in ["DELIV_QTY", "DELIV_PER", "TTL_TRD_QNTY", "CLOSE_PRICE", "HIGH_PRICE", "LOW_PRICE"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col].astype(str).str.strip(), errors="coerce")

df["DELIV_TURNOVER_CR"] = (df["DELIV_QTY"] * df["CLOSE_PRICE"]) / 1e7
df["SECTOR"] = df["SYMBOL"].map(sector_mapping).fillna("Others")

unique_dates = sorted(df["DATE1"].unique())
latest_market_date = unique_dates[-1].strftime("%d-%b-%Y")
recent_dates = unique_dates[-12:]
baseline_dates = unique_dates[:-12]

df_recent = df[df["DATE1"].isin(recent_dates)]
df_base = df[df["DATE1"].isin(baseline_dates)]

# =====================================================================
# 6. SECTOR ROTATION CALCULATION
# =====================================================================
base_total_cr = df_base["DELIV_TURNOVER_CR"].sum()
recent_total_cr = df_recent["DELIV_TURNOVER_CR"].sum()

base_sec = df_base.groupby("SECTOR")["DELIV_TURNOVER_CR"].sum().reset_index()
base_sec["3M Base Share (%)"] = (base_sec["DELIV_TURNOVER_CR"] / base_total_cr * 100).round(2)

recent_sec = df_recent.groupby("SECTOR")["DELIV_TURNOVER_CR"].sum().reset_index()
recent_sec["Recent 12D Share (%)"] = (recent_sec["DELIV_TURNOVER_CR"] / recent_total_cr * 100).round(2)

sector_rotation = pd.merge(
    recent_sec[["SECTOR", "Recent 12D Share (%)"]],
    base_sec[["SECTOR", "3M Base Share (%)"]],
    on="SECTOR", how="outer"
).fillna(0)

sector_rotation["Flow Shift (%)"] = (sector_rotation["Recent 12D Share (%)"] - sector_rotation["3M Base Share (%)"]).round(2)
sector_rotation["Flow Signal"] = sector_rotation["Flow Shift (%)"].apply(
    lambda x: "🟢 Heavy Inflow" if x >= 0.5 else ("🟢 Inflow" if x > 0 else ("🔴 Heavy Outflow" if x <= -0.5 else "🔴 Outflow"))
)
sector_rotation = sector_rotation.sort_values(by="Flow Shift (%)", ascending=False).reset_index(drop=True)
inflow_sectors = set(sector_rotation[sector_rotation["Flow Shift (%)"] > 0]["SECTOR"])

# =====================================================================
# 7. ACTIVE ACCUMULATION (TODAY)
# =====================================================================
base_stats = df_base.groupby("SYMBOL").agg(
    BASE_AVG_QTY=("DELIV_QTY", "mean"),
    BASE_AVG_DELIV_PER=("DELIV_PER", "mean")
).reset_index()

recent_stats = df_recent.groupby("SYMBOL").agg(
    RECENT_AVG_QTY=("DELIV_QTY", "mean"),
    CURRENT_PRICE=("CLOSE_PRICE", "last"),
    PRICE_CHG_PCT=("CLOSE_PRICE", lambda x: ((x.iloc[-1] - x.iloc[0]) / x.iloc[0]) * 100 if len(x) > 1 else 0),
    RECENT_AVG_DELIV_PER=("DELIV_PER", "mean"),
    RECENT_DELIV_CR=("DELIV_TURNOVER_CR", "sum"),
    SECTOR=("SECTOR", "first"),
    SWING_HIGH_KEY=("HIGH_PRICE", lambda x: x.iloc[:-1].max() if len(x) > 1 else x.iloc[-1]),
    SWING_LOW_KEY=("LOW_PRICE", "min")
).reset_index()

acc_df = pd.merge(recent_stats, base_stats, on="SYMBOL", how="inner")
acc_df["QTY_SPIKE_RATIO"] = (acc_df["RECENT_AVG_QTY"] / acc_df["BASE_AVG_QTY"]).round(2)

current_active = acc_df[
    (acc_df["QTY_SPIKE_RATIO"] >= 1.40) &
    (acc_df["PRICE_CHG_PCT"].between(-4.0, 10.0)) &
    (acc_df["RECENT_AVG_DELIV_PER"] >= 45.0) &
    (acc_df["RECENT_DELIV_CR"] >= 8.0)
].sort_values(by="QTY_SPIKE_RATIO", ascending=False).reset_index(drop=True)

# Process Active Table
tracking_status = []
first_detected_dates = []
holding_days_list = []
smc_status_list = []
bos_trigger_list = []
sl_level_list = []
sector_badges = []
final_actions = []
final_buy_prices = []
trailing_sl_list = []
signal_advice_list = []
news_remarks = []

for _, r in current_active.iterrows():
    sym = r["SYMBOL"]
    cmp = r["CURRENT_PRICE"]
    sw_high = r["SWING_HIGH_KEY"]
    sw_low = r["SWING_LOW_KEY"]
    sec = r["SECTOR"]

    if sym in prev_tracking_dict:
        tracking_status.append("⚡ ACTIVE HOLD")
        f_date = prev_tracking_dict[sym]
        first_detected_dates.append(f_date)
        try:
            d_obj = datetime.datetime.strptime(f_date, "%d-%b-%Y").date()
            h_days = (today - d_obj).days
        except Exception:
            h_days = 1
        holding_days_list.append(f"{max(h_days, 1)} Days")
    else:
        tracking_status.append("🟢 NEW ENTRY")
        first_detected_dates.append(latest_market_date)
        holding_days_list.append("1 Day")

    if cmp >= sw_high:
        smc_status_list.append("🚀 CONFIRMED BOS (BUY)")
    else:
        smc_status_list.append("⏳ ABSORPTION (WAIT)")
    bos_trigger_list.append(round(sw_high, 2))
    sl_level_list.append(round(sw_low, 2))

    if sec in inflow_sectors:
        sector_badges.append("✅ Inflow Aligned")
    else:
        sector_badges.append("⚠️ Sector Outflow")

    act = user_trade_actions.get(sym, "WATCHLIST")
    bp = user_buy_prices.get(sym, None)
    final_actions.append(act)
    final_buy_prices.append(bp if bp else "")

    if act == "ENTERED" and bp:
        pnl_pct = ((cmp - bp) / bp) * 100
        if pnl_pct >= 5.0:
            tsl = max(bp, sw_low)
            advice = f"🟢 HOLD & RIDE (+{pnl_pct:.1f}%) [SL @ Cost]"
        elif pnl_pct >= 12.0:
            tsl = max(bp * 1.05, sw_low)
            advice = f"🎯 BOOK 50% PROFIT (+{pnl_pct:.1f}%)"
        else:
            tsl = sw_low
            advice = f"⚡ POSITION ACTIVE ({pnl_pct:+.1f}%)"

        if cmp < tsl:
            advice = f"🔴 EXIT / SELL TODAY (SL Hit @ ₹{tsl:.2f})"
        trailing_sl_list.append(round(tsl, 2))
        signal_advice_list.append(advice)
    elif act == "DROP":
        trailing_sl_list.append("-")
        signal_advice_list.append("🚫 IGNORED / DROPPED")
    else:
        trailing_sl_list.append(round(sw_low, 2))
        if cmp >= sw_high:
            signal_advice_list.append("🚀 READY TO BUY (Enter Price)")
        else:
            signal_advice_list.append("⏳ WAIT FOR BOS TRIGGER")

    news_remarks.append(get_hinglish_catalyst(sym, sec))

current_active["Status"] = tracking_status
current_active["FirstDate"] = first_detected_dates
current_active["HoldingDays"] = holding_days_list
current_active["SMC_Status"] = smc_status_list
current_active["BOS_Trigger"] = bos_trigger_list
current_active["SL_Level"] = sl_level_list
current_active["Sector_Badge"] = sector_badges
current_active["Action"] = final_actions
current_active["BuyPrice"] = final_buy_prices
current_active["TrailingSL"] = trailing_sl_list
current_active["SignalAdvice"] = signal_advice_list
current_active["NewsRemark"] = news_remarks

clean_active = pd.DataFrame({
    "Report Date": latest_market_date,
    "Live Status": current_active["Status"],
    "Symbol": current_active["SYMBOL"],
    "Sector": current_active["SECTOR"],
    "Sector Alignment": current_active["Sector_Badge"],
    "SMC Structure": current_active["SMC_Status"],
    "Trade Action": current_active["Action"],
    "My Buy Price (Rs)": current_active["BuyPrice"],
    "CMP (Rs)": current_active["CURRENT_PRICE"].round(2),
    "BOS Trigger (Rs)": current_active["BOS_Trigger"],
    "Support / TSL (Rs)": current_active["TrailingSL"],
    "Trade Signal": current_active["SignalAdvice"],
    "Hinglish News & Catalyst Remark": current_active["NewsRemark"],
    "Volume Spurt": current_active["QTY_SPIKE_RATIO"],
    "Recent Deliv %": current_active["RECENT_AVG_DELIV_PER"].round(1),
    "Radar Age": current_active["HoldingDays"]
})

# =====================================================================
# 8. 1-WEEK HISTORICAL EXIT ENGINE (Hinglish Reasons)
# =====================================================================
exited_rows = []
active_today_syms = set(current_active["SYMBOL"])

if len(unique_dates) >= 18:
    w1_dates = unique_dates[-17:-5]
    w1_base = unique_dates[:-17]
    df_w1 = df[df["DATE1"].isin(w1_dates)]
    df_w1_base = df[df["DATE1"].isin(w1_base)]
    
    w1_base_agg = df_w1_base.groupby("SYMBOL")["DELIV_QTY"].mean().reset_index().rename(columns={"DELIV_QTY": "W1_BASE_QTY"})
    w1_agg = df_w1.groupby("SYMBOL").agg(
        W1_QTY=("DELIV_QTY", "mean"),
        W1_PRICE_CHG=("CLOSE_PRICE", lambda x: ((x.iloc[-1] - x.iloc[0]) / x.iloc[0]) * 100 if len(x) > 1 else 0),
        W1_DELIV_PER=("DELIV_PER", "mean"),
        W1_DELIV_CR=("DELIV_TURNOVER_CR", "sum")
    ).reset_index()
    
    w1_merged = pd.merge(w1_agg, w1_base_agg, on="SYMBOL", how="inner")
    w1_merged["W1_SPURT"] = (w1_merged["W1_QTY"] / w1_merged["W1_BASE_QTY"]).round(2)
    
    w1_qualified = set(w1_merged[
        (w1_merged["W1_SPURT"] >= 1.40) &
        (w1_merged["W1_PRICE_CHG"].between(-4.0, 10.0)) &
        (w1_merged["W1_DELIV_PER"] >= 45.0) &
        (w1_merged["W1_DELIV_CR"] >= 8.0)
    ]["SYMBOL"])
    
    hist_exits = list(w1_qualified - active_today_syms)
    
    for x_sym in hist_exits:
        match = acc_df[acc_df["SYMBOL"] == x_sym]
        if not match.empty:
            cmp_val = match["CURRENT_PRICE"].values[0]
            chg_val = match["PRICE_CHG_PCT"].values[0]
            spurt_val = match["QTY_SPIKE_RATIO"].values[0]
            
            if chg_val > 10.0:
                h_reason = f"🚀 Target Hit / Breakout Complete ({chg_val:+.1f}% move aa gaya)"
            elif chg_val < -4.0:
                h_reason = f"🛑 Support Broken / Structure Fail ({chg_val:+.1f}% drawdown)"
            elif spurt_val < 1.40:
                h_reason = f"📉 Volume Spurt khatam hua (Spurt {spurt_val:.2f}x par gir gaya)"
            else:
                h_reason = "⚠️ Delivery percentage threshold se niche chala gaya"
                
            exited_rows.append({
                "Symbol": x_sym,
                "Exit Date": unique_dates[-5].strftime("%d-%b-%Y"),
                "Current Price (Rs)": round(cmp_val, 2),
                "Recent Return (%)": round(chg_val, 2),
                "Last Spurt Ratio": spurt_val,
                "Hinglish Exit Reason": h_reason
            })

if prev_tracking_dict:
    for p_sym in set(prev_tracking_dict.keys()) - active_today_syms:
        if not any(r["Symbol"] == p_sym for r in exited_rows):
            match = acc_df[acc_df["SYMBOL"] == p_sym]
            if not match.empty:
                cmp_val = match["CURRENT_PRICE"].values[0]
                chg_val = match["PRICE_CHG_PCT"].values[0]
                spurt_val = match["QTY_SPIKE_RATIO"].values[0]
                exited_rows.append({
                    "Symbol": p_sym,
                    "Exit Date": latest_market_date,
                    "Current Price (Rs)": round(cmp_val, 2),
                    "Recent Return (%)": round(chg_val, 2),
                    "Last Spurt Ratio": spurt_val,
                    "Hinglish Exit Reason": "Latest session me criteria se bahar hua"
                })

exited_df = pd.DataFrame(exited_rows).sort_values(by="Recent Return (%)", ascending=False).reset_index(drop=True) if exited_rows else pd.DataFrame(
    columns=["Symbol", "Exit Date", "Current Price (Rs)", "Recent Return (%)", "Last Spurt Ratio", "Hinglish Exit Reason"]
)

# =====================================================================
# 9. PROFESSIONAL OPENPYXL STYLING & SAFE OVERWRITE
# =====================================================================
wb = openpyxl.Workbook()
wb.remove(wb.active)

C_NAVY = "0F172A"
C_WHITE = "FFFFFF"
C_BORDER = "CBD5E1"
C_NEW = "D1FAE5"
C_HOLD = "F0F9FF"
C_BOS = "DCFCE7"
C_WAIT = "FEF9C3"
C_CARD = "F8FAFC"

font_family = "Times New Roman"
title_font = Font(name=font_family, size=14, bold=True, color=C_NAVY)
sub_font = Font(name=font_family, size=11, italic=True, color="475569")
header_font = Font(name=font_family, size=13, bold=True, color=C_WHITE)
data_font = Font(name=font_family, size=12, bold=False, color="000000")
data_font_bold = Font(name=font_family, size=12, bold=True, color="000000")

thin_side = Side(style='thin', color=C_BORDER)
grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# --- SHEET 1: Top_Accumulation_Radar ---
ws1 = wb.create_sheet(title="Top_Accumulation_Radar")
ws1.views.sheetView[0].showGridLines = True

ws1.merge_cells("A1:P1")
ws1["A1"].value = "INSTITUTIONAL ACCUMULATION & SMC BREAKOUT RADAR"
ws1["A1"].font = title_font
ws1["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws1["A1"].fill = PatternFill(start_color=C_CARD, end_color=C_CARD, fill_type="solid")

ws1.merge_cells("A2:P2")
ws1["A2"].value = f"Data Synchronized As On: {latest_market_date} | Market Close Delivery & News Catalyst"
ws1["A2"].font = sub_font
ws1["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws1["A2"].fill = PatternFill(start_color=C_CARD, end_color=C_CARD, fill_type="solid")

ws1.row_dimensions[1].height = 26
ws1.row_dimensions[2].height = 20

headers_1 = list(clean_active.columns)
for col_idx, h_text in enumerate(headers_1, start=1):
    c = ws1.cell(row=3, column=col_idx, value=h_text)
    c.font = header_font
    c.fill = PatternFill(start_color=C_NAVY, end_color=C_NAVY, fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = grid_border
ws1.row_dimensions[3].height = 32

for r_idx, row_data in enumerate(clean_active.itertuples(index=False), start=4):
    status_val = str(row_data[1])
    act_val = str(row_data[6])
    smc_val = str(row_data[5])
    sig_val = str(row_data[11])
    is_new = "NEW" in status_val

    for c_idx, val in enumerate(row_data, start=1):
        cell = ws1.cell(row=r_idx, column=c_idx, value=val)
        cell.border = grid_border
        cell.font = data_font_bold if is_new or c_idx in [3, 7, 12] else data_font

        if act_val == "ENTERED":
            cell.fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
        elif is_new:
            cell.fill = PatternFill(start_color=C_NEW, end_color=C_NEW, fill_type="solid")
        else:
            cell.fill = PatternFill(start_color=C_HOLD, end_color=C_HOLD, fill_type="solid")

        if c_idx == 6:
            if "CONFIRMED" in smc_val:
                cell.fill = PatternFill(start_color=C_BOS, end_color=C_BOS, fill_type="solid")
            else:
                cell.fill = PatternFill(start_color=C_WAIT, end_color=C_WAIT, fill_type="solid")

        if c_idx == 12:
            if "EXIT" in sig_val:
                cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            elif "HOLD" in sig_val or "READY" in sig_val:
                cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

        col_name = headers_1[c_idx - 1]
        if "(Rs)" in col_name:
            if isinstance(val, (int, float)):
                cell.number_format = '₹#,##0.00'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif "%" in col_name:
            cell.number_format = '0.0"%"'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif "Spurt" in col_name:
            cell.number_format = '0.00"x"'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif "News" in col_name:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[r_idx].height = 26

# --- SHEET 2: Sector_Rotation ---
ws2 = wb.create_sheet(title="Sector_Rotation")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells("A1:E1")
ws2["A1"].value = f"NSE INSTITUTIONAL CAPITAL FLOW & SECTOR ROTATION ({latest_market_date})"
ws2["A1"].font = title_font
ws2["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws2["A1"].fill = PatternFill(start_color=C_CARD, end_color=C_CARD, fill_type="solid")
ws2.row_dimensions[1].height = 28

headers_2 = list(sector_rotation.columns)
for col_idx, h_text in enumerate(headers_2, start=1):
    c = ws2.cell(row=2, column=col_idx, value=h_text)
    c.font = header_font
    c.fill = PatternFill(start_color=C_NAVY, end_color=C_NAVY, fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = grid_border
ws2.row_dimensions[2].height = 30

for r_idx, row_data in enumerate(sector_rotation.itertuples(index=False), start=3):
    flow_signal = str(row_data[4])
    for c_idx, val in enumerate(row_data, start=1):
        cell = ws2.cell(row=r_idx, column=c_idx, value=val)
        cell.border = grid_border
        cell.font = data_font
        if "Inflow" in flow_signal:
            cell.fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
        else:
            cell.fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")

        col_name = headers_2[c_idx - 1]
        if "%" in col_name:
            cell.number_format = '0.00"%"'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif "SECTOR" in col_name:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[r_idx].height = 24

# --- SHEET 3: Exited_Stocks_Log ---
ws3 = wb.create_sheet(title="Exited_Stocks_Log")
ws3.views.sheetView[0].showGridLines = True

ws3.merge_cells("A1:F1")
ws3["A1"].value = "REMOVED / EXITED STOCKS AUDIT LOG (LAST 1-WEEK ANALYSIS)"
ws3["A1"].font = title_font
ws3["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws3["A1"].fill = PatternFill(start_color=C_CARD, end_color=C_CARD, fill_type="solid")
ws3.row_dimensions[1].height = 28

headers_3 = list(exited_df.columns)
for col_idx, h_text in enumerate(headers_3, start=1):
    c = ws3.cell(row=2, column=col_idx, value=h_text)
    c.font = header_font
    c.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = grid_border
ws3.row_dimensions[2].height = 30

for r_idx, row_data in enumerate(exited_df.itertuples(index=False), start=3):
    reason_txt = str(row_data[5])
    for c_idx, val in enumerate(row_data, start=1):
        cell = ws3.cell(row=r_idx, column=c_idx, value=val)
        cell.border = grid_border
        cell.font = data_font

        if "Target Hit" in reason_txt:
            cell.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
        else:
            cell.fill = PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid")

        col_name = headers_3[c_idx - 1]
        if "(Rs)" in col_name:
            cell.number_format = '₹#,##0.00'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif "(%)" in col_name:
            cell.number_format = '0.00"%"'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif "Spurt" in col_name:
            cell.number_format = '0.00"x"'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif "Reason" in col_name:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[r_idx].height = 24

# Auto-Fit Column Widths Across All Sheets
for sheet in [ws1, ws2, ws3]:
    for col_idx in range(1, sheet.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(3, sheet.max_row + 1):
            c_val = sheet.cell(row=row_idx, column=col_idx).value
            if c_val is not None:
                max_len = max(max_len, len(str(c_val)))
        sheet.column_dimensions[col_letter].width = max(max_len + 4, 16)

# SAFE OVERWRITE HANDLING (Agar Excel khuli ho tab bhi crash na ho)
try:
    wb.save(DASHBOARD_FILE)
    print(f"🎉 SUCCESS: Fresh Dashboard REPLACED & UPDATED at '{DASHBOARD_FILE}'!")
except PermissionError:
    print(f"⚠️ PERMISSION ERROR: 'SmartMoney_Live_Dashboard.xlsx' Excel me open hai!")
    print("👉 Kripya Excel file ko band (close) karein aur CMD me dobara run karein taaki wo replace ho sake.")