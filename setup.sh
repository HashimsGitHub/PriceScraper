#!/usr/bin/env bash
set -e

# 1. Create project directory structure
echo "Creating project directory structure..."
mkdir -p .github/workflows

# 2. Create requirements.txt
echo "Creating requirements.txt..."
cat << 'EOF' > requirements.txt
streamlit>=1.30.0
pandas>=2.0.0
openpyxl>=3.1.2
requests>=2.31.0
EOF

# 3. Create .github/workflows/main.yml
echo "Creating GitHub Actions workflow..."
cat << 'EOF' > .github/workflows/main.yml
name: Streamlit App Lint & Build Verification

on:
  push:
    branches: [ "main" ]
  pull_request:
    branches: [ "main" ]

jobs:
  build:
    runs-on: ubuntu-latest

    steps:
    - uses: actions/checkout@v3

    - name: Set up Python 3.10
      uses: actions/setup-python@v4
      with:
        python-version: "3.10"

    - name: Install Python Dependencies
      run: |
        python -m pip install --upgrade pip
        if [ -f requirements.txt ]; then pip install -r requirements.txt; fi

    - name: Check Streamlit Syntax
      run: |
        python -m py_compile app.py
EOF

# 4. Create app.py with SerpApi Google Shopping Integration
echo "Creating app.py..."
cat << 'EOF' > app.py
import streamlit as st
import pandas as pd
import requests
import io

st.set_page_config(
    page_title="Dell Core i9 Price Scraper",
    page_icon="💻",
    layout="wide"
)

def fetch_google_shopping_prices(query, target_count, api_key):
    """Fetch authentic US market listings using SerpApi Google Shopping."""
    url = "https://serpapi.com/search.json"
    params = {
        "engine": "google_shopping",
        "q": query,
        "location": "United States",
        "hl": "en",
        "gl": "us",
        "num": target_count,
        "api_key": api_key
    }
    
    try:
        response = requests.get(url, params=params, timeout=15)
        if response.status_code != 200:
            st.error(f"API Error ({response.status_code}): {response.text}")
            return []
            
        data = response.json()
        shopping_results = data.get("shopping_results", [])
        
        results = []
        for item in shopping_results[:target_count]:
            seller = item.get("source", "US Retailer")
            
            # Format clean seller website domain
            clean_domain = seller.lower().replace(" ", "").replace("'", "")
            seller_website = f"{clean_domain}.com" if not clean_domain.endswith(".com") else clean_domain

            results.append({
                "Seller Website": seller_website,
                "Seller Name": seller,
                "Item Description": item.get("title", "N/A"),
                "Item SKU": item.get("product_id", "N/A"),
                "Unit Price (USD)": item.get("price", "N/A")
            })
            
        return results

    except Exception as e:
        st.error(f"Failed to connect to SerpApi: {e}")
        return []

# --- UI Interface ---
st.title("💻 US Market Price Scraper: Dell Core i9 Desktops")
st.markdown("Scan authentic US merchants for **Dell Desktops (Core i9, 1TB NVMe, Windows 11)** via Google Shopping API.")

# Sidebar for API Key & Settings
with st.sidebar:
    st.header("🔑 API Configuration")
    api_key = st.text_input("Enter SerpApi Key:", type="password", help="Get a free key at https://serpapi.com")
    st.info("Using an official API prevents web scrapers from getting blocked by Cloudflare or captcha checks.")

col1, col2 = st.columns([2, 1])

with col1:
    num_items = st.slider("Select maximum number of prices to fetch:", min_value=5, max_value=50, value=15, step=5)
    search_query = st.text_input("Search Query:", value="DELL Desktop Core i9 1TB NVMe Windows 11")

with col2:
    st.markdown("### Active Parameters")
    st.caption("- **Target Region:** United States (`gl=us`)")
    st.caption("- **Engine:** `google_shopping`")
    st.caption("- **Output Format:** Excel (`.xlsx`)")

if st.button("🚀 Start Market Price Scan", type="primary"):
    if not api_key:
        st.warning("Please enter your SerpApi Key in the sidebar to proceed.")
    else:
        with st.spinner("Fetching authentic US merchant prices via SerpApi..."):
            scraped_data = fetch_google_shopping_prices(search_query, num_items, api_key)
            
        if scraped_data:
            df = pd.DataFrame(scraped_data)
            st.success(f"Successfully retrieved {len(df)} authentic listing(s)!")
            
            st.subheader("Data Preview")
            st.dataframe(df, use_container_width=True)
            
            # Excel Export
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Dell i9 Prices')
            
            st.download_button(
                label="📥 Download Results as Excel (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name="dell_i9_us_market_prices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No listings found for this search query. Try broadening your keywords.")
EOF

# 5. Create README.md
echo "Creating README.md..."
cat << 'EOF' > README.md
# Dell Core i9 Web Price Scraper

Streamlit web application that queries the US Google Shopping market using SerpApi for Dell Desktops with Intel Core i9, 1TB NVMe, and Windows 11 OS, and exports the data to Excel format.

## Setup Instructions

1. Install dependencies:
   ```bash
   pip install -r requirements.txt

EOF