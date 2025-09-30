import streamlit as st
import pandas as pd
from io import BytesIO
from rapidfuzz import process, fuzz
from datetime import datetime

st.set_page_config(page_title="Bank -> Tally Mapper", layout="wide")

def to_date_str(val):
    # normalize pandas datetime or string to dd-mm-YYYY
    try:
        dt = pd.to_datetime(val, dayfirst=False, errors='coerce')
        if pd.isna(dt):
            return ""
        return dt.strftime("%d-%m-%Y")
    except:
        return ""

def main():
    st.title("🏦 Bank Statement → Tally-style Mapper (DARSHH GOWDA)")

    st.markdown("""
    **Upload Bank Statement** (Excel `.xlsx`) with columns: `DATE`, `NARRITION`, `DEBIT`, `CREDIT`.  
    Optional: upload **Ledger Master** Excel (one column containing ledger names).
    """)

    col1, col2 = st.columns(2)
    with col1:
        bank_file = st.file_uploader("Upload Bank Statement Excel", type=["xlsx"])
    with col2:
        ledger_file = st.file_uploader("Upload Ledger Master Excel (optional)", type=["xlsx","csv"])

    st.write("---")
    # options
    fuzzy_thresh = st.slider("Fuzzy match threshold (ledger auto-match)", 50, 100, 80)
    give_bank_name = st.text_input("If you want to auto-fill Bank Name for blank bank side (type bank name, leave blank to skip):", value="")
    st.write("Note: Payment = Debit side (money out). Receipt = Credit side (money in).")

    if bank_file is None:
        st.info("Upload bank statement to proceed.")
        return

    # read bank statement
    try:
        df = pd.read_excel(bank_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading bank file: {e}")
        return

    # normalize column names to uppercase trimmed
    df.columns = [str(c).strip().upper() for c in df.columns]

    # required columns
    required = ["DATE", "NARRITION", "DEBIT", "CREDIT"]
    for r in required:
        if r not in df.columns:
            st.error(f"Required column '{r}' not found in uploaded bank statement. Columns found: {list(df.columns)}")
            return

    # coerce numeric
    df["DEBIT"] = pd.to_numeric(df["DEBIT"], errors="coerce").fillna(0)
    df["CREDIT"] = pd.to_numeric(df["CREDIT"], errors="coerce").fillna(0)

    # show preview
    st.subheader("Uploaded bank statement (preview)")
    st.dataframe(df.head(10))

    # build repeated narrations list
    counts = df["NARRITION"].astype(str).str.strip().str.upper().value_counts()
    repeated = counts[counts > 1].index.tolist()

    st.subheader("Repeated narrations (count > 1) — enter manual replacements if you want")
    st.write("Only items shown below — you can type mapped ledger/label for each.")
    # store replacements
    narration_map = {}
    for item in repeated:
        # show original case examples: find first matching original string to display
        sample = df[df["NARRITION"].astype(str).str.upper() == item]["NARRITION"].iloc[0]
        user_inp = st.text_input(f"Replace '{sample}' (repeated) ->", key=f"rep_{item}")
        if user_inp.strip():
            narration_map[item] = user_inp.strip()

    # ledger master load
    ledger_list = []
    if ledger_file is not None:
        try:
            if str(ledger_file.name).lower().endswith(".csv"):
                df_led = pd.read_csv(ledger_file)
            else:
                df_led = pd.read_excel(ledger_file, engine="openpyxl")
            # pick first column as ledger names
            if df_led.shape[1] >= 1:
                ledger_list = df_led.iloc[:,0].astype(str).str.strip().tolist()
                ledger_list = [x for x in ledger_list if x]
                st.success(f"Loaded {len(ledger_list)} ledger names from master file.")
            else:
                st.warning("Ledger file has no columns.")
        except Exception as e:
            st.error(f"Error reading ledger master: {e}")

    # preview of debit<=58 detection count
    cnt_charges = int((df["DEBIT"] <= 58).sum())
    st.info(f"Rows with DEBIT <= 58 will be auto-mapped to 'BANK CHARGES' (count = {cnt_charges}).")

    # button to apply mapping
    if st.button("Apply mapping & Generate Tally-style Output"):
        working = df.copy()
        # initialize mapped ledger column as empty
        working["MAPPED_LEDGER"] = ""

        # 1) apply manual repeated replacements (exact match case-insensitive)
        for orig_upper, mapped in narration_map.items():
            mask = working["NARRITION"].astype(str).str.upper() == orig_upper
            working.loc[mask, "MAPPED_LEDGER"] = mapped

        # 2) Debit <= 58 => BANK CHARGES (overwrite or set)
        mask_charges = working["DEBIT"] <= 58
        working.loc[mask_charges, "MAPPED_LEDGER"] = "BANK CHARGES"

        # 3) Fuzzy match remaining using ledger master (if provided)
        if ledger_list:
            # we'll try to match only rows where MAPPED_LEDGER is empty
            need_match_idx = working[working["MAPPED_LEDGER"] == ""].index
            narrs_to_match = working.loc[need_match_idx, "NARRITION"].astype(str).tolist()
            # for speed, use process.extractOne per narration
            for idx in need_match_idx:
                text = str(working.at[idx, "NARRITION"])
                if not text.strip():
                    continue
                best = process.extractOne(text, ledger_list, scorer=fuzz.token_set_ratio)
                if best:
                    match_name, score, _ = best
                    if score >= fuzzy_thresh:
                        working.at[idx, "MAPPED_LEDGER"] = match_name

        # 4) Now build Tally-style columns
        out_rows = []
        for i, row in working.iterrows():
            date_raw = row["DATE"]
            date_str = to_date_str(date_raw)
            # amount: pick debit if >0 else credit
            debit = float(row["DEBIT"]) if pd.notna(row["DEBIT"]) else 0.0
            credit = float(row["CREDIT"]) if pd.notna(row["CREDIT"]) else 0.0
            amount = debit if debit > 0 else credit

            # voucher type
            if debit > 0 and credit == 0:
                voucher_type = "PAYMENT"
            elif credit > 0 and debit == 0:
                voucher_type = "RECEIPT"
            else:
                # both zero or both >0 (edge) -> decide by sign: prefer PAYMENT if debit>credit else RECEIPT
                voucher_type = "PAYMENT" if debit >= credit else "RECEIPT"

            mapped = row.get("MAPPED_LEDGER", "")
            narration_text = str(row["NARRITION"])

            # determine BY/DR and TO/CR according to rules:
            by_dr = ""
            to_cr = ""

            bank_side_filled = False
            # if voucher is RECEIPT: BY/DR = bank name, TO/CR = mapped ledger
            if voucher_type == "RECEIPT":
                if give_bank_name.strip():
                    by_dr = give_bank_name.strip()
                    bank_side_filled = True
                if mapped:
                    to_cr = mapped
                else:
                    to_cr = "SUSPENSE"
            else:  # PAYMENT
                if mapped:
                    by_dr = mapped
                else:
                    by_dr = "SUSPENSE"
                if give_bank_name.strip():
                    to_cr = give_bank_name.strip()
                    bank_side_filled = True
                else:
                    to_cr = "SUSPENSE"

            # voucher no blank
            voucher_no = ""
            # day from date
            day_val = ""
            try:
                if date_str:
                    day_val = datetime.strptime(date_str, "%d-%m-%Y").strftime("%d")
                else:
                    day_val = ""
            except:
                day_val = ""

            out_rows.append({
                "DATE": date_str,
                "VOUCHER NO.": voucher_no,
                "BY / DR": by_dr,
                "TO / CR": to_cr,
                "AMOUNT": amount,
                "NARRATION": narration_text,
                "VOUCHER TYPE": voucher_type,
                "DAY": day_val
            })

        out_df = pd.DataFrame(out_rows, columns=["DATE","VOUCHER NO.","BY / DR","TO / CR","AMOUNT","NARRATION","VOUCHER TYPE","DAY"])

        st.success("Mapping complete — preview below.")
        st.dataframe(out_df.head(20))

        # provide download
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Tally_Import")
        st.download_button("⬇️ Download Final Excel", data=buffer.getvalue(), file_name="tally_mapped_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # also show count summary
        st.write("### Summary")
        st.write(f"Total rows: {len(out_df)}")
        st.write("Voucher type counts:")
        st.write(out_df["VOUCHER TYPE"].value_counts())

if __name__ == "__main__":
    main()
