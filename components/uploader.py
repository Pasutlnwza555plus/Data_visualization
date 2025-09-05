import pandas as pd
import streamlit as st
from services.session import SessionStateManager, SessionStateEnum

class ExcelUploader:
    def __init__(self, session: SessionStateManager):
        self.session = session

    def upload(self, title: str, session_state: SessionStateEnum):
        key_name = title.lower().replace(" ", "_")
        uploaded_file = st.file_uploader(f"Upload {title}", type=["xlsx"], key=str(key_name))

        if uploaded_file:
            df = pd.read_excel(uploaded_file)
            self.session[session_state] = df

            st.success(f"{title} Uploaded")