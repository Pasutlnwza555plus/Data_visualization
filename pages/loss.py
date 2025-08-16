import streamlit as st
import pandas as pd


# region Base Analyzer for Loss
class LossAnalyzer:
    def __init__(self):
        self.df_ref = st.session_state.get("core_eol_reference_sheet")
        self.df_raw_data = st.session_state.get("raw_eol_data")

    # ------- Utilities -------
    @staticmethod
    def is_castable_to_float(x) -> bool:
        try:
            float(x)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def countDay(df_ref: pd.DataFrame):
        days = (len(df_ref.columns) - 11) / 4
        return int(days)

    @staticmethod
    def isDiffError(row):
        color = [''] * len(row)
        if float(row["Loss current - Loss EOL"]) >= 2:
            color = ['background-color: #ff4d4d; color: white'] * len(row)
        elif row["Remark"].strip() != "":
            color = ['background-color: #d6b346; color: white'] * len(row)
        return color

    @staticmethod
    def draw_color_legend():
        st.markdown("""
            <div style='display: flex; justify-content: center; align-items: center; gap: 16px; margin-bottom: 1rem'>
                <div style='display: flex; justify-content: center; align-items: center; gap: 8px'>
                    <div style='background-color: #ff4d4d; width: 24px; height: 24px; border-radius: 8px;'></div>
                    <div style='text-align: center; color: #ff4d4d; font-size: 24px; font-weight: bold;'>
                        EOL NOT OK 
                    </div>
                </div>
                <div style='display: flex; justify-content: center; align-items: center; gap: 8px'>
                    <div style='background-color: #d6b346; width: 24px; height: 24px; border-radius: 8px;'></div>
                    <div style='text-align: center; color: #d6b346; font-size: 24px; font-weight: bold;'>
                        Fiber break occurs
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)

    # Abstract placeholder (children must override)
    def process(self):
        raise NotImplementedError("Each analyzer must implement its own process()")


# region Analyzer for EOL
class EOLAnalyzer(LossAnalyzer):

    def extract_eol_ref(self, df_ref: pd.DataFrame) -> pd.DataFrame:
        df_eol_ref = pd.DataFrame()
        eol_ref_columns = df_ref.columns[df_ref.iloc[0] == "EOL(dB)"]
        eol_ref_columns_float = pd.to_numeric(df_ref[eol_ref_columns[0]], downcast="float", errors="coerce")
        df_eol_ref["Link Name"] = df_ref['140.1'].iloc[1:]
        df_eol_ref["EOL(dB)"] = eol_ref_columns_float
        return df_eol_ref

    def extract_raw_data(self, df_raw_data: pd.DataFrame) -> pd.DataFrame:
        df_raw_data.columns = df_raw_data.columns.str.strip()
        df_atten = pd.DataFrame()
        source_port_col = df_raw_data["Source Port"]
        sink_port_col = df_raw_data["Sink Port"]
        df_atten["Link Name"] = source_port_col + "_" + sink_port_col
        df_atten["Current Attenuation(dB)"] = df_raw_data["Optical Attenuation (dB)"]
        df_atten["Remark"] = df_atten["Current Attenuation(dB)"].apply(
            lambda x: "" if self.is_castable_to_float(x) else "Fiber Break"
        )
        return df_atten

    def calculate_eol_diff(self, df_eol: pd.DataFrame) -> pd.DataFrame:
        df_eol_diff = df_eol.copy()
        current_atten_col = pd.to_numeric(df_eol["Current Attenuation(dB)"], downcast="float", errors="coerce")
        eol_ref_col = pd.to_numeric(df_eol["EOL(dB)"], downcast="float", errors="coerce")
        calculated_diff = current_atten_col - eol_ref_col - 1
        df_eol_diff["Loss current - Loss EOL"] = calculated_diff
        ordered_cols = ["Link Name", "EOL(dB)", "Current Attenuation(dB)", "Loss current - Loss EOL", "Remark"]
        return df_eol_diff[ordered_cols]

    def process(self):
        if self.df_ref is not None and self.df_raw_data is not None:
            df_eol_ref = self.extract_eol_ref(self.df_ref)
            df_atten = self.extract_raw_data(self.df_raw_data)
            joined_df = df_eol_ref.join(df_atten.set_index("Link Name"), on="Link Name")
            df_result = self.calculate_eol_diff(joined_df)
            st.dataframe(df_result.style.apply(self.isDiffError, axis=1), hide_index=True)
            self.draw_color_legend()


# region Analyzer for Core
class CoreAnalyzer(LossAnalyzer):

    def process(self):
        if self.df_raw_data is not None:
            # In future thou mayest define core-specific logic here
            st.write("Core Analyzer processing not yet implemented.")
            st.dataframe(self.df_raw_data.head())
