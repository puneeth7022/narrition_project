import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
from rapidfuzz import process, fuzz
from datetime import datetime

st.set_page_config(page_title="Bank -> Tally Mapper", layout="wide")

# 🔹 Helper: Date normalize
def to_date_str(val):
    try:
        dt = pd.to_datetime(val, dayfirst=False, errors='coerce')
        if pd.isna(dt):
            return ""
        return dt.strftime("%d-%m-%Y")
    except:
        return ""

import pdfplumber
import pandas as pd

def parse_canara_pdf(file):
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table:
                    # Skip empty rows
                    if any(cell and str(cell).strip() for cell in row):
                        # Keep only first 8 columns if extra present
                        if len(row) > 8:
                            row = row[:8]
                        rows.append(row)

    # Define 8 headers (standard Canara Bank format)
    headers = [
        "TRANS_DATE", "VALUE_DATE", "BRANCH", "REF_CHQNO",
        "DESCRIPTION", "WITHDRAWS", "DEPOSIT", "BALANCE"
    ]
    
    df = pd.DataFrame(rows[1:], columns=headers)

    # Clean numeric columns
    for col in ["WITHDRAWS", "DEPOSIT", "BALANCE"]:
        df[col] = df[col].astype(str).str.replace(",", "").str.strip()
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Final output format
    df_out = pd.DataFrame({
        "DATE": df["TRANS_DATE"],
        "NARRITION": df["DESCRIPTION"],
        "DEBIT": df["WITHDRAWS"],
        "CREDIT": df["DEPOSIT"]
    })
    
    return df_out


# 🔹 Main App
def main():
    st.title("🏦 Bank Statement → Tally-style Mapper (Darshan Edition)")

    st.markdown("""
    **Upload Bank Statement**  
    - Excel `.xlsx` with columns: `DATE`, `NARRITION`, `DEBIT`, `CREDIT`  
    - OR Canara Bank PDF statement (digital PDF only)
    """)

    file = st.file_uploader("Upload Bank Statement (Excel or PDF)", type=["xlsx","pdf"])
    ledger_file = st.file_uploader("Upload Ledger Master Excel (optional)", type=["xlsx","csv"])

    fuzzy_thresh = st.slider("Fuzzy match threshold (ledger auto-match)", 50, 100, 80)
    give_bank_name = st.text_input("Bank Name (auto-fill in output):", value="")

    if file is None:
        st.info("Upload a bank statement to proceed.")
        return

    # Detect file type
    if str(file.name).lower().endswith(".pdf"):
        try:
            df = parse_canara_pdf(file)
            st.success("✅ Parsed Canara Bank PDF successfully.")
        except Exception as e:
            st.error(f"Error parsing PDF: {e}")
            return
    else:
        try:
            df = pd.read_excel(file, engine="openpyxl")
        except Exception as e:
            st.error(f"Error reading Excel: {e}")
            return

    df.columns = [str(c).strip().upper() for c in df.columns]
    required = ["DATE","NARRITION","DEBIT","CREDIT"]
    for r in required:
        if r not in df.columns:
            st.error(f"Missing column '{r}' in statement. Found: {list(df.columns)}")
            return

    df["DEBIT"] = pd.to_numeric(df["DEBIT"], errors="coerce").fillna(0)
    df["CREDIT"] = pd.to_numeric(df["CREDIT"], errors="coerce").fillna(0)

    st.subheader("Uploaded statement (preview)")
    st.dataframe(df.head(10))

    # 🔹 Mapping Logic
    counts = df["NARRITION"].astype(str).str.strip().str.upper().value_counts()
    repeated = counts[counts > 1].index.tolist()

    st.subheader("Repeated Narrations (manual replacement)")
    narration_map = {}
    for item in repeated:
        sample = df[df["NARRITION"].astype(str).str.upper() == item]["NARRITION"].iloc[0]
        user_inp = st.text_input(f"Replace '{sample}' →", key=f"rep_{item}")
        if user_inp.strip():
            narration_map[item] = user_inp.strip()

    ledger_list = []
    if ledger_file is not None:
        try:
            if str(ledger_file.name).lower().endswith(".csv"):
                df_led = pd.read_csv(ledger_file)
            else:
                df_led = pd.read_excel(ledger_file, engine="openpyxl")
            if df_led.shape[1] >= 1:
                ledger_list = df_led.iloc[:,0].astype(str).str.strip().tolist()
                ledger_list = [x for x in ledger_list if x]
                st.success(f"Loaded {len(ledger_list)} ledger names.")
        except Exception as e:
            st.error(f"Error reading ledger file: {e}")

    cnt_charges = int((df["DEBIT"] <= 58).sum())
    st.info(f"Rows with DEBIT <= 58 auto-mapped to 'BANK CHARGES' (count={cnt_charges}).")

    if st.button("Generate Tally Output"):
        working = df.copy()
        working["MAPPED_LEDGER"] = ""

        # Manual replacements
        for orig_upper, mapped in narration_map.items():
            mask = working["NARRITION"].astype(str).str.upper() == orig_upper
            working.loc[mask, "MAPPED_LEDGER"] = mapped

        # Debit <= 58 → BANK CHARGES
        mask_charges = working["DEBIT"] <= 58
        working.loc[mask_charges, "MAPPED_LEDGER"] = "BANK CHARGES"

        # Fuzzy match with Ledger Master
        if ledger_list:
            need_match_idx = working[working["MAPPED_LEDGER"] == ""].index
            for idx in need_match_idx:
                text = str(working.at[idx,"NARRITION"])
                best = process.extractOne(text, ledger_list, scorer=fuzz.token_set_ratio)
                if best and best[1] >= fuzzy_thresh:
                    working.at[idx,"MAPPED_LEDGER"] = best[0]

        # Build Tally-style output
        out_rows = []
        for _, row in working.iterrows():
            date_str = to_date_str(row["DATE"])
            debit, credit = float(row["DEBIT"]), float(row["CREDIT"])
            amount = debit if debit > 0 else credit

            if debit > 0 and credit == 0:
                vtype = "PAYMENT"
            elif credit > 0 and debit == 0:
                vtype = "RECEIPT"
            else:
                vtype = "PAYMENT" if debit >= credit else "RECEIPT"

            mapped = row["MAPPED_LEDGER"]
            narration = str(row["NARRITION"])
            by_dr, to_cr = "", ""

            if vtype == "RECEIPT":
                by_dr = give_bank_name if give_bank_name else "BANK"
                to_cr = mapped if mapped else "SUSPENSE"
            else:
                by_dr = mapped if mapped else "SUSPENSE"
                to_cr = give_bank_name if give_bank_name else "BANK"

            out_rows.append({
                "DATE": date_str,
                "VOUCHER NO.": "",
                "BY / DR": by_dr,
                "TO / CR": to_cr,
                "AMOUNT": amount,
                "NARRATION": narration,
                "VOUCHER TYPE": vtype,
                "DAY": date_str[:2] if date_str else ""
            })

        out_df = pd.DataFrame(out_rows, columns=["DATE","VOUCHER NO.","BY / DR","TO / CR","AMOUNT","NARRATION","VOUCHER TYPE","DAY"])
        st.success("✅ Mapping complete.")
        st.dataframe(out_df.head(20))

        # Download
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Tally_Import")
        st.download_button("⬇️ Download Final Excel", buffer.getvalue(), "tally_mapped_output.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()



