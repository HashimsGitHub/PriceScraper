# Product Price Scraper

Streamlit app that searches Google Shopping through SerpApi and prepares selected
listings for the supplied PrestaShop product-entry workbook.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Keep `PrestaShop_Product_Entry_Template.xlsx` beside `app.py`.
3. Add `SERPAPI_KEY` to Streamlit Secrets, or enter the key in the app's sidebar.
4. Start the app:
   ```bash
   streamlit run app.py
   ```

## Workflow

1. Enter a product search and choose the PrestaShop tax rule/state for its results.
2. Repeat searches as needed; results and selections remain available in the current
   Streamlit session.
3. Review the listings and select products. Google Shopping fills the available product
   title, Google product ID, merchant, link, and price. Product type and publish status
   default to the template's `Standard product` and `Offline` values; choose a tax
   rule/state for each search. Add or correct other details when known.
4. Download **PrestaShop_Product_Entry_Data.xlsx** to get the selected products appended
   to a copy of the original template. Fields not returned by Shopping or entered by
   you, including description, category, quantity, and weight, remain blank. The app
   shows data-quality warnings but does not block export when these fields are blank.
   Other sheets and template formatting are kept.
5. Download **PrestaShop_Search_Results.xlsx** for the complete Shopping results from
   every search in the session. It includes the raw result fields returned by SerpApi
   and a search-run summary.

Dropdown-backed fields must match values on the template's `Lists` sheet. The category
list in the supplied template may need to be expanded with your store's categories.
Google Shopping prices are prefilled as candidate retail prices; verify each one and
adjust it to the correct tax-exclusive price before importing. Google product IDs are
candidate references, not verified manufacturer SKUs. Confirm product identifiers and
any blank template fields against the product/merchant before importing to PrestaShop.
