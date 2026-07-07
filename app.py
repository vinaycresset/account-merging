"""
CCT ↔ Addepar reconciliation app.

Upload a CCT holdings file and an Addepar file. The app normalizes the name
columns (uppercase + strip a leading "CCT - " prefix), then updates the Addepar
"Value as of 3/31" column with the Market Value from the CCT file by matching:

    CCT "Name"    <-> Addepar "Holding Account"
    CCT "Account" <-> Addepar "Top Level Legal Entity"

Reconciliation rules:
    * Matched rows          -> Value as of 3/31 = summed Market Value from CCT
    * In Addepar, not CCT   -> highlighted RED
    * In CCT, not Addepar   -> appended with Owner Id / Entity ID blank,
                               highlighted ORANGE
"""

import io

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Column names
# ---------------------------------------------------------------------------
CCT_NAME = "Name"
CCT_ACCOUNT = "Account"
CCT_MARKET_VALUE = "Market Value"

ADP_HOLDING_ACCOUNT = "Holding Account"
ADP_OWNER_ID = "Owner Id"
ADP_ENTITY_ID = "Entity ID"
ADP_VALUE = "Value as of 3/31"
ADP_TLLE = "Top Level Legal Entity"

STATUS_COL = "Match Status"
STATUS_MATCHED = "Matched"
STATUS_ADP_ONLY = "In Addepar, not CCT"
STATUS_CCT_ONLY = "In CCT, not Addepar"

RED = "background-color: #f8b4b4"     # in Addepar but not CCT
ORANGE = "background-color: #fcd9a5"  # in CCT but not Addepar


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def read_any(uploaded_file) -> pd.DataFrame:
    """Read an uploaded CSV or Excel file into a DataFrame."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def normalize_name(series: pd.Series) -> pd.Series:
    """Uppercase, trim, and strip a leading 'CCT - ' prefix."""
    s = series.astype("string").str.strip().str.upper()
    # Remove a leading "CCT - " (allowing flexible spacing around the dash).
    s = s.str.replace(r"^CCT\s*-\s*", "", regex=True).str.strip()
    return s


def to_number(series: pd.Series) -> pd.Series:
    """Coerce currency-like strings ('$1,234.56') to floats."""
    if pd.api.types.is_numeric_dtype(series):
        return series
    cleaned = (
        series.astype("string")
        .str.replace(r"[\$,]", "", regex=True)
        .str.replace(r"^\((.*)\)$", r"-\1", regex=True)  # (123) -> -123
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce")


def require_columns(df: pd.DataFrame, cols, label) -> list:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        st.error(f"{label} file is missing columns: {missing}")
        st.write(f"Columns found in {label}: {list(df.columns)}")
    return missing


# ---------------------------------------------------------------------------
# Core reconciliation
# ---------------------------------------------------------------------------
def reconcile(cct: pd.DataFrame, addepar: pd.DataFrame) -> pd.DataFrame:
    cct = cct.copy()
    addepar = addepar.copy()

    # Normalize the join-key name columns in both frames.
    cct[CCT_NAME] = normalize_name(cct[CCT_NAME])
    cct[CCT_ACCOUNT] = normalize_name(cct[CCT_ACCOUNT])
    addepar[ADP_HOLDING_ACCOUNT] = normalize_name(addepar[ADP_HOLDING_ACCOUNT])
    addepar[ADP_TLLE] = normalize_name(addepar[ADP_TLLE])

    cct[CCT_MARKET_VALUE] = to_number(cct[CCT_MARKET_VALUE])

    # Aggregate CCT Market Value per (Name, Account) — multiple securities
    # under one account/entity are summed.
    cct_agg = (
        cct.groupby([CCT_NAME, CCT_ACCOUNT], dropna=False, as_index=False)[
            CCT_MARKET_VALUE
        ]
        .sum()
        .rename(
            columns={
                CCT_NAME: ADP_HOLDING_ACCOUNT,
                CCT_ACCOUNT: ADP_TLLE,
                CCT_MARKET_VALUE: "_cct_value",
            }
        )
    )

    keys = [ADP_HOLDING_ACCOUNT, ADP_TLLE]
    merged = addepar.merge(cct_agg, on=keys, how="outer", indicator=True)

    # Determine status from the merge indicator.
    #   both       -> matched
    #   left_only  -> Addepar row with no CCT match (RED)
    #   right_only -> CCT row with no Addepar match (ORANGE, new row)
    status = merged["_merge"].map(
        {
            "both": STATUS_MATCHED,
            "left_only": STATUS_ADP_ONLY,
            "right_only": STATUS_CCT_ONLY,
        }
    )

    # Update Value as of 3/31 for matched rows and populate new CCT-only rows.
    is_match = merged["_merge"] == "both"
    is_cct_only = merged["_merge"] == "right_only"

    merged.loc[is_match, ADP_VALUE] = merged.loc[is_match, "_cct_value"]
    merged.loc[is_cct_only, ADP_VALUE] = merged.loc[is_cct_only, "_cct_value"]

    # CCT-only rows have no Owner Id / Entity ID.
    for col in (ADP_OWNER_ID, ADP_ENTITY_ID):
        if col in merged.columns:
            merged.loc[is_cct_only, col] = pd.NA

    merged[STATUS_COL] = status
    merged = merged.drop(columns=["_merge", "_cct_value"])

    # Keep Addepar's original column order, with Match Status at the end.
    ordered = [c for c in addepar.columns if c in merged.columns]
    ordered += [c for c in merged.columns if c not in ordered]
    return merged[ordered]


def style_rows(df: pd.DataFrame):
    """Return a Styler that colors rows by match status."""
    def color(row):
        if row[STATUS_COL] == STATUS_ADP_ONLY:
            return [RED] * len(row)
        if row[STATUS_COL] == STATUS_CCT_ONLY:
            return [ORANGE] * len(row)
        return [""] * len(row)

    return df.style.apply(color, axis=1)


def to_excel_bytes(styler) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        styler.to_excel(writer, index=False, sheet_name="Reconciled")
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.set_page_config(page_title="CCT ↔ Addepar Reconciliation", layout="wide")
st.title("CCT ↔ Addepar Reconciliation")

st.markdown(
    "Upload both files, then the Addepar **Value as of 3/31** is updated with "
    "the CCT **Market Value**.\n\n"
    "- **Red** — in Addepar but not CCT\n"
    "- **Orange** — in CCT but not Addepar (added with blank Owner Id / Entity ID)"
)

col1, col2 = st.columns(2)
with col1:
    cct_file = st.file_uploader(
        "CCT", type=["xlsx", "xls", "csv"], key="cct"
    )
with col2:
    addepar_file = st.file_uploader(
        "Addepar", type=["xlsx", "xls", "csv"], key="addepar"
    )

if cct_file and addepar_file:
    cct_df = read_any(cct_file)
    addepar_df = read_any(addepar_file)

    missing = require_columns(
        cct_df, [CCT_NAME, CCT_ACCOUNT, CCT_MARKET_VALUE], "CCT"
    )
    missing += require_columns(
        addepar_df,
        [ADP_HOLDING_ACCOUNT, ADP_TLLE, ADP_VALUE, ADP_OWNER_ID, ADP_ENTITY_ID],
        "Addepar",
    )

    if not missing:
        result = reconcile(cct_df, addepar_df)

        # Summary counts.
        counts = result[STATUS_COL].value_counts()
        c1, c2, c3 = st.columns(3)
        c1.metric("Matched", int(counts.get(STATUS_MATCHED, 0)))
        c2.metric("In Addepar, not CCT", int(counts.get(STATUS_ADP_ONLY, 0)))
        c3.metric("In CCT, not Addepar", int(counts.get(STATUS_CCT_ONLY, 0)))

        styler = style_rows(result)
        st.dataframe(styler, use_container_width=True)

        st.download_button(
            "Download reconciled Excel",
            data=to_excel_bytes(styler),
            file_name="reconciled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload both a CCT file and an Addepar file to begin.")
