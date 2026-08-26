#!/bin/bash
cd "$(dirname "$0")"
export STREAMLIT_CONFIG_DIR="$(pwd)/.streamlit"
export STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
source .venv/bin/activate
streamlit run app.py --server.headless true --server.port 8501
