import streamlit as st
import gspread
from google.oauth2 import service_account

# Create a connection object.
credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
    ],
)
gc = gspread.authorize(credentials)

sheet_url = st.secrets["private_gsheets_url"]
sh = gc.open(sheet_url)
st.write(sh.sheet1.get('A1'))
