import pandas as pd
import numpy as np
from datetime import datetime

INPUT_FILE = r"C:\Users\湯川悦英\Downloads\坂井竜世 (1).xlsx"
OUTPUT_FILE = r"C:\Users\湯川悦英\Downloads\坂井竜世_案件インポート.csv"
SHEET_NAME = "2025坂井竜世"

# Month definitions: (column_letter_base_0_index_offset_from_H, month_number, year)
# H=col7, I=col8, J=col9, K=col10 → August
# Columns A=0, B=1, ..., H=7
MONTH_COLS = [
    # (見込_col_idx, 売上_col_idx, 費用_col_idx, 利益_col_idx, month_name, year, month_num)
    (7,  8,  9,  10, "8月",  2025, 8),   # H,I,J,K
    (11, 12, 13, 14, "9月",  2025, 9),   # L,M,N,O
    (15, 16, 17, 18, "10月", 2025, 10),  # P,Q,R,S
    (19, 20, 21, 22, "11月", 2025, 11),  # T,U,V,W
    (23, 24, 25, 26, "12月", 2025, 12),  # X,Y,Z,AA
    (27, 28, 29, 30, "1月",  2026, 1),   # AB,AC,AD,AE
    (31, 32, 33, 34, "2月",  2026, 2),   # AF,AG,AH,AI
    (35, 36, 37, 38, "3月",  2026, 3),   # AJ,AK,AL,AM
    (39, 40, 41, 42, "4月",  2026, 4),   # AN,AO,AP,AQ
    (43, 44, 45, 46, "5月",  2026, 5),   # AR,AS,AT,AU
    (47, 48, 49, 50, "6月",  2026, 6),   # AV,AW,AX,AY
    (51, 52, 53, 54, "7月",  2026, 7),   # AZ,BA,BB,BC
]

def to_number(val):
    """Convert a value to a number, return 0 if not possible."""
    if pd.isna(val):
        return 0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0

def format_date(val):
    """Format a date value as YYYY-MM-DD."""
    if pd.isna(val):
        return ""
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%Y-%m-%d")
    try:
        return pd.to_datetime(val).strftime("%Y-%m-%d")
    except Exception:
        return str(val)

def map_tantosha(name):
    """Map short name to full name."""
    if pd.isna(name):
        return ""
    name = str(name).strip()
    if name == "坂井":
        return "坂井竜世"
    return name

def main():
    print(f"Reading: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME, skiprows=12, header=0)

    total_rows = len(df)
    print(f"Total rows loaded (after header): {total_rows}")

    output_rows = []
    skipped_shisshu = 0
    skipped_no_profit = 0
    case_counter = 0

    for idx, row in df.iterrows():
        # Col A=0: 入力日, B=1: 担当者, C=2: 決定確度, D=3: 企業名, E=4: 内容
        # F=5: 次回連絡日/決定月, G=6: 次の行動

        kakunin = row.iloc[2] if len(row) > 2 else None
        kakunin_str = str(kakunin).strip() if not pd.isna(kakunin) else ""

        # Skip if 決定確度 is '失注' or empty
        if kakunin_str == "" or kakunin_str == "nan" or kakunin_str == "失注":
            skipped_shisshu += 1
            continue

        # Find max profit month
        best_profit = 0
        best_month_info = None
        best_sales = 0
        best_cost = 0

        for (miko_col, uriage_col, hiyo_col, rieki_col, month_name, year, month_num) in MONTH_COLS:
            if rieki_col >= len(row):
                continue
            rieki = to_number(row.iloc[rieki_col])
            if rieki > best_profit:
                best_profit = rieki
                best_sales = to_number(row.iloc[uriage_col]) if uriage_col < len(row) else 0
                best_cost = to_number(row.iloc[hiyo_col]) if hiyo_col < len(row) else 0
                best_month_info = (year, month_num)

        # Skip if no FY2025 month has non-zero 利益
        if best_month_info is None or best_profit == 0:
            skipped_no_profit += 1
            continue

        case_counter += 1
        case_id = f"SAKAI-{case_counter:03d}"

        input_date = row.iloc[0] if len(row) > 0 else None
        tantosha_raw = row.iloc[1] if len(row) > 1 else None
        kigyo_name = row.iloc[3] if len(row) > 3 else ""
        naiyou = row.iloc[4] if len(row) > 4 else ""
        next_action = row.iloc[6] if len(row) > 6 else ""

        toroku_bi = format_date(input_date)
        tantosha = map_tantosha(tantosha_raw)
        kaisha_mei = "" if pd.isna(kigyo_name) else str(kigyo_name).strip()
        shozai_mei = "" if pd.isna(naiyou) else str(naiyou).strip()
        memo = "" if pd.isna(next_action) else str(next_action).strip()

        uriage_yotei_tsuki = f"{best_month_info[0]}-{best_month_info[1]:02d}"

        # 入金ステータス
        if kakunin_str == "売上":
            nyukin_status = "入金済み"
        else:
            nyukin_status = "未入金"

        output_rows.append({
            "案件ID": case_id,
            "登録日": toroku_bi,
            "担当者": tantosha,
            "顧客ID": "",
            "会社名": kaisha_mei,
            "商材名": shozai_mei,
            "フェーズ": "",
            "確度ランク": kakunin_str,
            "売上（単価）": int(best_sales) if best_sales == int(best_sales) else best_sales,
            "費用（単価）": int(best_cost) if best_cost == int(best_cost) else best_cost,
            "コース数": 1,
            "件数": 1,
            "月数": 1,
            "売上予定額": int(best_sales) if best_sales == int(best_sales) else best_sales,
            "費用（合計）": int(best_cost) if best_cost == int(best_cost) else best_cost,
            "粗利": int(best_profit) if best_profit == int(best_profit) else best_profit,
            "インセンティブ": 0,
            "売上予定月": uriage_yotei_tsuki,
            "入金ステータス": nyukin_status,
            "入金確認日": "",
            "メモ": memo,
            "引継担当者": "",
            "引継日": "",
        })

    print(f"\n--- Summary ---")
    print(f"Total rows read:          {total_rows}")
    print(f"Skipped (失注 or empty):  {skipped_shisshu}")
    print(f"Skipped (no 利益):        {skipped_no_profit}")
    print(f"Rows passed filter:       {case_counter}")

    if output_rows:
        out_df = pd.DataFrame(output_rows, columns=[
            "案件ID", "登録日", "担当者", "顧客ID", "会社名", "商材名",
            "フェーズ", "確度ランク", "売上（単価）", "費用（単価）",
            "コース数", "件数", "月数", "売上予定額", "費用（合計）",
            "粗利", "インセンティブ", "売上予定月", "入金ステータス",
            "入金確認日", "メモ", "引継担当者", "引継日"
        ])
        out_df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
        print(f"\nSaved: {OUTPUT_FILE}")
        print(f"Output rows: {len(out_df)}")
    else:
        print("No rows matched the filter criteria. CSV not created.")

if __name__ == "__main__":
    main()
