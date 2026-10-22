#!/usr/bin/env bash
set -euo pipefail

if [[ ! -f requirements.txt || ! -f app.py ]]; then
    echo "Run setup.sh from the PriceScraper project directory." >&2
    exit 1
fi

if [[ ! -f PrestaShop_Product_Entry_Template.xlsx ]]; then
    echo "PrestaShop_Product_Entry_Template.xlsx must be beside app.py." >&2
    exit 1
fi

echo "Project files are ready. Install dependencies with:"
echo "  python -m pip install -r requirements.txt"
echo "Then start the app with:"
echo "  streamlit run app.py"
