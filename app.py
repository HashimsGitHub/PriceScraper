import streamlit as st
import pandas as pd
import requests
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

            # Clean raw price string into numeric float for proper Excel formatting
            raw_price = item.get("price", "0")
            price_numeric = float(str(raw_price).replace("$", "").replace(",", "").strip()) if raw_price else 0.0

            results.append({
                "Seller Website": seller_website,
                "Seller Name": seller,
                "Item Description": item.get("title", "N/A"),
                "Item SKU": item.get("product_id", "N/A"),
                "Unit Price (USD)": price_numeric
            })
            
        return results

    except Exception as e:
        st.error(f"Failed to connect to SerpApi: {e}")
        return []

def generate_formatted_excel(df):
    """Generates a neatly formatted Excel file with custom styles, headers, and column widths."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dell i9 Market Prices')
        
        workbook = writer.book
        worksheet = writer.sheets['Dell i9 Market Prices']
        
        # --- Styles Definition ---
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # Dark Steel Blue
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        
        zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")   # Very Light Blue/Gray
        data_font = Font(name="Calibri", size=10)
        
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        right_align = Alignment(horizontal="right", vertical="center")

        # --- Format Headers ---
        for col_num in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
            
        # Set Header Row Height
        worksheet.row_dimensions[1].height = 28

        # --- Format Data Rows ---
        for row_num in range(2, len(df) + 2):
            worksheet.row_dimensions[row_num].height = 22
            is_even = (row_num % 2 == 0)
            
            for col_num in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                cell.font = data_font
                cell.border = thin_border
                
                # Apply Zebra Striping
                if is_even:
                    cell.fill = zebra_fill
                    
                col_name = df.columns[col_num - 1]
                
                # Column Specific Alignment & Number Formatting
                if col_name == "Unit Price (USD)":
                    cell.number_format = '"$"#,##0.00'
                    cell.alignment = right_align
                elif col_name in ["Seller Website", "Item SKU"]:
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align

        # --- Auto-fit Column Widths ---
        for col_idx, col in enumerate(df.columns, 1):
            max_len = max(
                df[col].astype(str).map(len).max() if not df.empty else 0,
                len(str(col))
            )
            col_letter = get_column_letter(col_idx)
            # Add padding for breathing room
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 50)

    return output.getvalue()

# --- UI Interface ---
st.title("💻 US Market Price Scraper: Dell Core i9 Desktops")
st.markdown("Scan authentic US merchants for **Dell Desktops (Core i9, 1TB NVMe, Windows 11)** via Google Shopping API.")

# Secrets Management & Key Retrieval
api_key = st.secrets.get("SERPAPI_KEY")

with st.sidebar:
    st.header("⚙️ Settings & Configuration")
    if api_key:
        st.success("✅ SerpAPI Key loaded from Secrets")
    else:
        api_key = st.text_input("Enter SerpApi Key:", type="password", help="Get a free key at https://serpapi.com")
        st.warning("⚠️ No secret key detected in Streamlit Settings.")

col1, col2 = st.columns([2, 1])

with col1:
    num_items = st.slider("Select maximum number of prices to fetch:", min_value=5, max_value=50, value=15, step=5)
    search_query = st.text_input("Search Query:", value="DELL Desktop Core i9 1TB NVMe Windows 11")

with col2:
    st.markdown("### Active Parameters")
    st.caption("- **Target Region:** United States (`gl=us`)")
    st.caption("- **Engine:** `google_shopping`")
    st.caption("- **Export Styling:** Dark Steel Blue & Zebra Striping")

if st.button("🚀 Start Market Price Scan", type="primary"):
    if not api_key:
        st.error("Please configure `SERPAPI_KEY` in Streamlit Secrets or enter it manually in the sidebar.")
    else:
        with st.spinner("Fetching authentic US merchant prices via SerpApi..."):
            scraped_data = fetch_google_shopping_prices(search_query, num_items, api_key)
            
        if scraped_data:
            df = pd.DataFrame(scraped_data)
            st.success(f"Successfully retrieved {len(df)} authentic listing(s)!")
            
            st.subheader("Data Preview")
            st.dataframe(
                df.style.format({"Unit Price (USD)": "${:,.2f}"}),
                use_container_width=True
            )
            
            # Generate styled Excel file
            excel_bytes = generate_formatted_excel(df)
            
            st.download_button(
                label="📥 Download Styled Excel Report (.xlsx)",
                data=excel_bytes,
                file_name="dell_i9_us_market_prices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No listings found for this search query. Try broadening your keywords.")