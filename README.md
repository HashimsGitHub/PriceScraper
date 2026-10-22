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
3. Review the listings, select products, and complete required template fields such as
   SKU, description, category, quantity, weight, price (tax excluded), tax rule/state,
   and publish status. Product type and publish status default to the template's
   `Standard product` and `Offline` values.
4. Download **PrestaShop_Product_Entry_Data.xlsx** to get the selected products appended
   to a copy of the original template. Other sheets and template formatting are kept.
5. Download **PrestaShop_Search_Results.xlsx** for the complete Shopping results from
   every search in the session. It includes the raw result fields returned by SerpApi
   and a search-run summary.

Dropdown-backed fields must match values on the template's `Lists` sheet. The category
list in the supplied template may need to be expanded with your store's categories.
Google Shopping prices are prefilled as candidate retail prices; verify that each one
is appropriate for the template's tax-exclusive price field before importing.
