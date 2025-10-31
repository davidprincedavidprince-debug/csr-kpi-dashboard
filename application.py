import os
import streamlit.web.cli as stcli

if __name__ == '__main__':
    os.system("streamlit run quant_scoring_reckoner.py --server.port 8080 --server.address 0.0.0.0")
