import gspread
import streamlit as st
from oauth2client.service_account import ServiceAccountCredentials

def get_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = st.secrets["gcp_service_account"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
    client = gspread.authorize(credentials)
    return client.open("Life Settlement Policies").sheet1
