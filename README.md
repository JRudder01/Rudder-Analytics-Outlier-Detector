# Housing Sale Outlier Highlighter

A small Streamlit app for reviewing home-sale outliers using two pasted Excel columns:

- Close Price
- Living Area

The app keeps every listing, highlights statistical outliers in blue, previews the result in the browser, and exports a one-sheet Excel workbook.

## Output formatting

- Close Price: commas, no currency symbol, no decimals
- Living Area: no commas, no decimals
- Data font: Aptos Narrow, 11 pt, unbolded
- Outliers: blue fill only
- Workbook sheet: `Highlighted_Output`

## Files

- `streamlit_app.py` — Streamlit interface
- `outlier_engine.py` — parsing, outlier analysis, and Excel export logic
- `requirements.txt` — Python dependencies
- `.gitignore` — standard files GitHub should ignore

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy with GitHub + Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload all files in this folder to the root of the repository.
3. Commit the files to the `main` branch.
4. Open Streamlit Community Cloud and create/deploy an app from the GitHub repository.
5. Select `streamlit_app.py` as the app entrypoint.
6. Deploy.

No secrets or API keys are required.

## Normal workflow

1. In Excel, copy the Close Price and Living Area columns together.
2. Paste them into the app's input box.
3. Click **Analyze outliers**.
4. Review the highlighted output table.
5. Click **Download highlighted Excel file**.

The default rule is **Either method**, matching the final Colab setup: a listing is highlighted if it is flagged by either the robust price-vs-living-area model or the price-per-square-foot IQR test.
