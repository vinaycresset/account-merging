"""Generate fake CCT and Addepar files to test the reconciliation app.

Creates cct_sample.xlsx and addepar_sample.xlsx in this folder. The data is
built to exercise every path in the app:

  * A trust with two securities  -> matched, market values summed
  * A "CCT - " prefixed name      -> prefix stripped, still matches
  * A lowercase name              -> uppercased, still matches
  * "Ghost Account"               -> only in Addepar  -> RED
  * "Orphan In CCT"               -> only in CCT       -> ORANGE
  * A leading title row above the header (tests header auto-detection)
"""

import pandas as pd

# --- CCT file -------------------------------------------------------------
cct = pd.DataFrame(
    [
        # Smith trust: two securities under the same account -> summed to 15,000
        ["CCT - Smith Family Trust", "Cresset Trust Co", "Apple Inc", "AAPL",
         "037833100", 50, 200.00, "$10,000.00"],
        ["CCT - Smith Family Trust", "Cresset Trust Co", "Microsoft", "MSFT",
         "594918104", 10, 500.00, "$5,000.00"],
        # lowercase name -> uppercased and matched
        ["jones holdings llc", "Cresset Partners", "US Treasury", "T",
         "912810TX6", 25, 1000.00, "$25,000.00"],
        # only in CCT -> ORANGE, Owner Id / Entity ID left blank
        ["Orphan In CCT", "Cresset Advisors", "Cash", "USD",
         "", 1, 999.00, "$999.00"],
    ],
    columns=[
        "Name", "Account", "Security Description", "Symbol", "CUSIP",
        "Quantity", "Price", "Market Value",
    ],
)

# --- Addepar file ---------------------------------------------------------
addepar = pd.DataFrame(
    [
        ["Smith Family Trust", "OWN-100", "ENT-100", 0, "Cresset Trust Co"],
        ["Jones Holdings LLC", "OWN-200", "ENT-200", 0, "Cresset Partners"],
        # only in Addepar -> RED
        ["Ghost Account", "OWN-300", "ENT-300", 12345, "Cresset Advisors"],
    ],
    columns=[
        "Holding Account", "Owner Id", "Entity ID",
        "Value as of 3/31", "Top Level Legal Entity",
    ],
)


def write_with_title(df, path, title):
    """Write an .xlsx with a title row above the header (like real exports)."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # Title on row 1, blank row 2, header + data starting row 3.
        pd.DataFrame([[title]]).to_excel(
            writer, sheet_name="Sheet1", index=False, header=False, startrow=0
        )
        df.to_excel(
            writer, sheet_name="Sheet1", index=False, startrow=2
        )


write_with_title(cct, "cct_sample.xlsx", "CCT Holdings Export — As of 3/31/2026")
write_with_title(
    addepar, "addepar_sample.xlsx", "Addepar Positions — As of 3/31/2026"
)
print("Wrote cct_sample.xlsx and addepar_sample.xlsx")
