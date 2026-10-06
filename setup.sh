#!/usr/bin/env bash
set -e

# 1. Create directory structure
echo "Creating project directory structure..."
mkdir -p .github/workflows

# 2. Create packages.txt
echo "Creating packages.txt..."
cat << 'EOF' > packages.txt
chromium
chromium-driver
EOF

# 3. Create requirements.txt
echo "Creating requirements.txt..."
cat << 'EOF' > requirements.txt
streamlit>=1.30.0
selenium>=4.15.0
chromedriver-autoinstaller>=0.6.2
pandas>=2.0.0
openpyxl>=3.1.2
requests>=2.31.0
beautifulsoup4>=4.12.0
EOF

# 4. Create .github/workflows/main.yml
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

    - name: Install System Dependencies (Chrome & Driver)
      run: |
        sudo apt-get update
        sudo apt-get install -y chromium-browser chromium-chromedriver

    - name: Install Python Dependencies
      run: |
        python -m pip install --upgrade pip
        if [ -f requirements.txt ]; then pip install -r requirements.txt; fi

    - name: Check Streamlit Syntax
      run: |
        python -m py_compile app.py
EOF

# 5. Create app.py with Streamlit Cloud permission fix
echo "Creating app.py..."
cat << 'EOF' > app.py
import os
import shutil
import time
import io
from urllib.parse import quote_plus
import streamlit as st
import pandas as pd
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup

st.set_page_config(
    page_title="Dell Core i9 Price Scraper",
    page_icon="💻",
    layout="wide"
)

def get_driver():
    """Setup Headless Chrome compatible with Streamlit Community Cloud and local environments."""
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    # Check for system Chromium binary (Streamlit Cloud via packages.txt)
    chromium_path = shutil.which("chromium") or shutil.which("chromium-browser")
    if chromium_path:
        options.binary_location = chromium_path

    # Check for system ChromeDriver binary
    driver_path = shutil.which("chromedriver") or "/usr/bin/chromedriver"
    if os.path.exists(driver_path):
        service = Service(executable_path=driver_path)
        return webdriver.Chrome(service=service, options=options)

    # Fallback for local development (download to user-writable /tmp folder)
    try:
        installed_path = chromedriver_autoinstaller.install(path="/tmp")
        service = Service(executable_path=installed_path)
        return webdriver.Chrome(service=service, options=options)
    except Exception as e:
        st.error(f"Failed to initialize ChromeDriver: {e}")
        raise e

def scrape_dell_official(target_count):
    results = []
    url = "https://www.dell.com/en-us/search/dell%20desktop%20i9%201tb%20nvme%20windows%2011"
    driver = get_driver()
    try:
        driver.get(url)
        time.sleep(4)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        products = soup.find_all('article', class_=lambda c: c and 'ps-stack' in c) or \
                   soup.find_all('div', class_=lambda c: c and 'ps-title' in str(c))
        
        for item in products:
            if len(results) >= target_count:
                break
            title_elem = item.find('h3') or item.find('a')
            price_elem = item.find('span', class_=lambda c: c and 'price' in str(c).lower())
            
            if title_elem and price_elem:
                title = title_elem.get_text(strip=True)
                price_text = price_elem.get_text(strip=True)
                if "i9" in title.lower():
                    results.append({
                        "Seller Website": "dell.com",
                        "Seller Name": "Dell Official US Store",
                        "Item Description": title,
                        "Item SKU": item.get('data-sku', 'N/A'),
                        "Unit Price (USD)": price_text
                    })
    except Exception as e:
        st.error(f"Error fetching Dell US: {e}")
    finally:
        driver.quit()
    return results

def scrape_google_shopping(target_count):
    results = []
    search_query = quote_plus("DELL Desktop Core i9 1TB NVMe Windows 11")
    url = f"https://www.google.com/search?tbm=shop&q={search_query}&hl=en&gl=us"
    driver = get_driver()
    try:
        driver.get(url)
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        cards = soup.find_all('div', class_=lambda c: c and ('sh-dgr__content' in str(c) or 'sh-np__click-target' in str(c)))
        
        for card in cards:
            if len(results) >= target_count:
                break
            title_elem = card.find('h3') or card.find('h4')
            price_elem = card.find('span', class_=lambda c: c and '$' in str(c))
            seller_elem = card.find('div', class_=lambda c: c and ('merchant' in str(c).lower() or 'seller' in str(c).lower()))
            
            if title_elem and price_elem:
                title = title_elem.get_text(strip=True)
                price = price_elem.get_text(strip=True)
                seller = seller_elem.get_text(strip=True) if seller_elem else "US Retailer"
                
                results.append({
                    "Seller Website": f"{seller.lower().replace(' ', '')}.com",
                    "Seller Name": seller,
                    "Item Description": title,
                    "Item SKU": "N/A",
                    "Unit Price (USD)": price
                })
    except Exception as e:
        st.error(f"Error fetching aggregated market prices: {e}")
    finally:
        driver.quit()
    return results

st.title("💻 US Market Price Scraper: Dell Core i9 Desktops")
st.markdown("Scan authentic US merchants for **Dell Desktops (Core i9, 1TB NVMe, Windows 11)** and export the dataset directly to Excel.")

col1, col2 = st.columns([2, 1])

with col1:
    num_items = st.slider("Select maximum number of listings to scrape:", min_value=5, max_value=50, value=15, step=5)

with col2:
    st.markdown("### Search Parameters")
    st.caption("- **Brand:** Dell")
    st.caption("- **CPU:** Intel Core i9")
    st.caption("- **Storage:** 1TB NVMe SSD")
    st.caption("- **OS:** Windows 11")

if st.button("🚀 Start Market Price Scan", type="primary"):
    with st.spinner("Initializing Headless Web Driver and Scanning US Retailers..."):
        scraped_data = []
        
        dell_results = scrape_dell_official(num_items)
        scraped_data.extend(dell_results)
        
        remaining_count = num_items - len(scraped_data)
        if remaining_count > 0:
            market_results = scrape_google_shopping(remaining_count)
            scraped_data.extend(market_results)
            
    if scraped_data:
        df = pd.DataFrame(scraped_data).head(num_items)
        st.success(f"Successfully scraped {len(df)} authentic listing(s)!")
        
        st.subheader("Data Preview")
        st.dataframe(df, use_container_width=True)
        
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
        st.warning("No listings matched the exact query. Try running the scan again or adjusting slider parameters.")
EOF

# 6. Create README.md
echo "Creating README.md..."
cat << 'EOF' > README.md
# Dell Core i9 Web Price Scraper

Streamlit web application that scans the US market for Dell Desktops with Intel Core i9, 1TB NVMe, and Windows 11 OS, exporting data directly to Excel format.

## Setup Instructions

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
  
EOF