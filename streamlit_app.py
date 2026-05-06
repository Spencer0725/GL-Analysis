import io
import pandas as pd
import streamlit as st

# ==============================
# CONFIG – adjust if your layout changes
# ==============================

RECON_SHEET_NAME = "113520000"
DEC_SHEET_NAME   = "11352000_Dec"
AMOUNT_COL_NAME  = "        Amount in LC"   # must match Dec file exactly

# Row indices for key recon lines (0-based)
ROW_ACCRUAL_PRICING = 9       # "Accrual via Pricing Condition"
ROW_RECLASS_PPV     = 10      # "Accrual reclass from PPV (Hi-P/Trisa)"
ROW_RECLASS_WEB     = 11      # "Accrual Reclass (WEB)"
ROW_OTHERS_TRUEUP   = 12      # "Others (true up)"
ROW_INV_RAVISA      = 16      # "Invoices - Ravisa"
ROW_ENDING_BAL      = 27      # "Ending Balance"
ROW_GL_BAL_SAP      = 29      # "GL Balance in SAP"
ROW_CHECK           = 30      # "check --> must be zero"

# Column indices (0-based)
COL_YTD = 6        # YTD column
COL_NOV = 17       # November column
COL_DEC = 18       # December column


# ==============================
# CORE LOGIC
# ==============================

def compute_dec_breakdown(dec_df: pd.DataFrame) -> dict:
    """
    Compute December amounts by recon line based on mapping rules.
    Returns a dict with:
      accrual_pricing, reclass_ppv, reclass_web, others_trueup, inv_ravisa, total
    """
    # Ensure numeric
    dec_df[AMOUNT_COL_NAME] = pd.to_numeric(
        dec_df[AMOUNT_COL_NAME], errors="coerce"
    ).fillna(0)

    # Split by Type
    df_we = dec_df[dec_df["Type"] == "WE"]
    df_kr = dec_df[dec_df["Type"] == "KR"]
    df_sa = dec_df[dec_df["Type"] == "SA"]

    total_we = df_we[AMOUNT_COL_NAME].sum()
    total_kr = df_kr[AMOUNT_COL_NAME].sum()

    # SA by OffsetAcct
    sa_group = df_sa.groupby("OffsetAcct")[AMOUNT_COL_NAME].sum()

    def sa_amount(key):
        return float(sa_group.get(key, 0.0))

    # Mapping from our discussion:
    sa_220812000 = sa_amount(220812000)      # Import duties → Accrual via Pricing
    sa_220811000 = sa_amount(220811000)      # Trisa Tariff / XI-China → PPV reclass
    sa_220810002 = sa_amount(220810002)      # Customs WEB Reclass → WEB reclass
    sa_ccusp1c   = sa_amount("CCUSP1C")      # Duties recl → Others (true up)

    accrual_pricing = total_we + sa_220812000
    reclass_ppv     = sa_220811000
    reclass_web     = sa_220810002
    others_trueup   = sa_ccusp1c
    inv_ravisa      = total_kr   # All KR = Invoices – Ravisa

    total = (
        accrual_pricing
        + reclass_ppv
        + reclass_web
        + others_trueup
        + inv_ravisa
    )

    return {
        "accrual_pricing": accrual_pricing,
        "reclass_ppv": reclass_ppv,
        "reclass_web": reclass_web,
        "others_trueup": others_trueup,
        "inv_ravisa": inv_ravisa,
        "total": total,
    }


def update_recon_with_dec(recon_df: pd.DataFrame, dec_df: pd.DataFrame) -> pd.DataFrame:
    """
    Update the reconciliation DataFrame with December data.
    """
    dec = compute_dec_breakdown(dec_df)

    def add_to_ytd(row_idx, amount):
        current = recon_df.iloc[row_idx, COL_YTD]
        try:
            base = float(current)
        except (TypeError, ValueError):
            base = 0.0
        recon_df.iloc[row_idx, COL_YTD] = base + amount

    # December column values
    recon_df.iloc[ROW_ACCRUAL_PRICING, COL_DEC] = dec["accrual_pricing"]
    recon_df.iloc[ROW_RECLASS_PPV,     COL_DEC] = dec["reclass_ppv"]
    recon_df.iloc[ROW_RECLASS_WEB,     COL_DEC] = dec["reclass_web"]
    recon_df.iloc[ROW_OTHERS_TRUEUP,   COL_DEC] = dec["others_trueup"]
    recon_df.iloc[ROW_INV_RAVISA,      COL_DEC] = dec["inv_ravisa"]

    # Update YTD
    add_to_ytd(ROW_ACCRUAL_PRICING, dec["accrual_pricing"])
    add_to_ytd(ROW_RECLASS_PPV,     dec["reclass_ppv"])
    add_to_ytd(ROW_RECLASS_WEB,     dec["reclass_web"])
    add_to_ytd(ROW_OTHERS_TRUEUP,   dec["others_trueup"])
    add_to_ytd(ROW_INV_RAVISA,      dec["inv_ravisa"])

    # Ending balance Dec
    ending_nov = recon_df.iloc[ROW_ENDING_BAL, COL_NOV]
    try:
        ending_nov_val = float(ending_nov)
    except (TypeError, ValueError):
        raise ValueError("Could not read November ending balance from recon file.")

    ending_dec = ending_nov_val + dec["total"]

    recon_df.iloc[ROW_ENDING_BAL, COL_DEC] = ending_dec
    recon_df.iloc[ROW_ENDING_BAL, COL_YTD] = ending_dec  # YE balance

    # GL Balance in SAP and check
    recon_df.iloc[ROW_GL_BAL_SAP, COL_DEC] = ending_dec
    recon_df.iloc[ROW_CHECK, COL_DEC] = (
        float(recon_df.iloc[ROW_ENDING_BAL, COL_DEC])
        - float(recon_df.iloc[ROW_GL_BAL_SAP, COL_DEC])
    )

    return recon_df


# ==============================
# STREAMLIT UI
# ==============================

st.title("POC – Automated GL Reconciliation for 113520000 (Dec)")

st.write("**Step 1:** Upload the existing reconciliation file (Jan–Nov completed).")
recon_file = st.file_uploader("Reconciliation file (.xlsx)", type=["xlsx"])

st.write("**Step 2:** Upload the December SAP extract for account 113520000.")
dec_file = st.file_uploader("December source data (.xlsx)", type=["xlsx"])

if recon_file and dec_file:
    if st.button("Generate December Reconciliation"):
        try:
            # Use openpyxl explicitly for reading
            recon_df = pd.read_excel(
                recon_file,
                sheet_name=RECON_SHEET_NAME,
                header=None,
                engine="openpyxl",
            )
            dec_df = pd.read_excel(
                dec_file,
                sheet_name=DEC_SHEET_NAME,
                engine="openpyxl",
            )

            updated_df = update_recon_with_dec(recon_df, dec_df)

            # Write updated workbook to in-memory buffer
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                updated_df.to_excel(
                    writer,
                    sheet_name=RECON_SHEET_NAME,
                    header=False,
                    index=False,
                )
            output.seek(0)

            st.success("December reconciliation created successfully.")
            st.download_button(
                label="Download updated reconciliation",
                data=output,
                file_name="US01_113520000_with_Dec.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except ImportError as e:
            st.error(
                "The 'openpyxl' package is not installed in this environment. "
                "Please add 'openpyxl' to requirements.txt and redeploy."
            )
        except Exception as e:
            st.error(f"Error while generating reconciliation: {e}")
else:
    st.info("Upload both the recon file and the December source data to proceed.")

    
