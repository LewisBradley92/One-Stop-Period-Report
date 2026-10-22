# One Stop Periodic Reporting App

A Streamlit app for preparing the One Stop Test Purchase Report workbook for each reporting period.

## Files required each run

1. The latest One Stop report workbook. For the first run, use `One-Stop Test Purchase Report P7 - LIVE.xlsx`. For later periods, upload the most recently generated workbook so the YTD tab retains its history.
2. The current audit export CSV.
3. The current One Stop Master Schedule.
4. The Tesco period calendar used as the calendar reference.

## Rules

- Only One Stop audits are considered when `client_name` is present.
- The audit `site_internal_id` must be in the Master Schedule's `NS ID` / `SITEID` column. Sites not on the list are excluded automatically.
- Only Alcohol, E-Cig and Cigarettes are included.
- The product category must match the site's assignment in the selected `Period N` column of the Master Schedule (`Cigarettes` is matched to `Tobacco`; `E-Cig` is matched to `E-Cigs`).
- The audit date must fall within the selected One Stop period.
- One Stop periods start Sunday and end Saturday. The Tesco calendar's Monday week-start dates are shifted to derive One Stop boundaries.
- Aborted audits are excluded. Fail audits are not excluded merely because they failed.
- `This month` is replaced each run.
- `YTD` is appended to, and duplicate audit IDs (`internal_id`) already in YTD are skipped.
- The existing column positions and workbook tabs are retained. Formula cells are preserved and Excel recalculation is requested when the output is opened.

## Run locally

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Put `app.py`, `report_builder.py`, and `requirements.txt` in a GitHub repository.
2. In Streamlit Community Cloud, create an app pointing to `app.py`.
3. Upload the four input files in the app each reporting period.
4. Download the generated workbook and use it as the report workbook input for the next period.

## Important validation note

The app uses the existing workbook's common question columns (product type/details, ID checks, staff interaction and receipt fields). The current source export includes additional product-specific questions that do not have dedicated labelled columns in the supplied template; those extra questions are not added as new columns because the agreed requirement is to preserve the existing report layout. Review the generated workbook in Excel and confirm that the template's summary formulas recalculate as expected before relying on it for client delivery.

## Client-facing export

The app produces two files for the selected period:

- `One-Stop Test Purchase Report P<N> - LIVE.xlsx` — the full internal workbook.
- `One-Stop Test Purchase Report P<N>.xlsx` — the client-facing workbook containing only `Summary Data`, `Store Performance`, `Regional Performance`, `North Region Performance`, `Central Region Performance`, `South Region Performance`, and `Product Performance`. Formula cells in this client copy are converted to calculated values.

`Summary Data` is trimmed dynamically to row 7 plus the number of current-period audits, so it has no trailing audit rows without data. The client-facing values-only export uses LibreOffice to recalculate formulas before converting them to values. `packages.txt` installs LibreOffice Calc on supported Streamlit Community Cloud deployments.
