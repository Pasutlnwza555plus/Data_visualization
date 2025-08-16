import math
import streamlit as st
import pandas as pd


# region Base Analyzer for Loss
class LossAnalyzer:
    def __init__(
        self, 
        df_ref: pd.DataFrame | None = None, 
        df_raw_data: pd.DataFrame | None = None
    ):
        self.df_ref = df_ref
        self.df_raw_data = df_raw_data

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
        status = ""
        if float(row["Loss current - Loss EOL"]) >= 2:
            status = "error"
        elif row["Remark"].strip() != "":
            status = "flapping"

        color_style = LossAnalyzer.getColor(status)
        return [color_style] * len(row)
    
    @staticmethod
    def getColor(status: str) -> str:
        color = ''
        if status == "error":
            color = 'background-color: #ff4d4d; color: white;'
        elif status == "flapping":
            color = 'background-color: #d6b346; color: white;'
        
        return color

    @staticmethod
    def draw_color_legend():
        st.markdown("""
            <div style='display: flex; justify-content: center; align-items: center; gap: 16px; margin-bottom: 1rem'>
                <div style='display: flex; justify-content: center; align-items: center; gap: 8px'>
                    <div style='background-color: #ff4d4d; width: 24px; height: 24px; border-radius: 8px;'></div>
                    <div style='text-align: center; color: #ff4d4d; font-size: 24px; font-weight: bold;'>
                        EOL not OK 
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

    # Base data extraction
    @staticmethod
    def extract_eol_ref(df_ref: pd.DataFrame) -> pd.DataFrame:
        df_eol_ref = pd.DataFrame()
        eol_ref_columns = df_ref.columns[df_ref.iloc[0] == "EOL(dB)"]
        eol_ref_columns_float = pd.to_numeric(df_ref[eol_ref_columns[0]], downcast="float", errors="coerce")

        df_eol_ref["Link Name"] = df_ref['140.1'].iloc[1:]
        df_eol_ref["EOL(dB)"] = eol_ref_columns_float

        return df_eol_ref

    # Abstract placeholder (children must override)
    def process(self):
        raise NotImplementedError("Each analyzer must implement its own process()")


# region Analyzer for EOL
class EOLAnalyzer(LossAnalyzer):
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
    
    def build_result_df(self):
        if self.df_ref is not None and self.df_raw_data is not None:
            df_eol_ref: pd.DataFrame = self.extract_eol_ref(self.df_ref)
            df_atten: pd.DataFrame   = self.extract_raw_data(self.df_raw_data)

            joined_df = df_eol_ref.join(df_atten.set_index("Link Name"), on="Link Name")
            df_result = self.calculate_eol_diff(joined_df)

        return df_result
    
    def get_me_names(self, df_result: pd.DataFrame) -> list[str]:
        link_names = df_result["Link Name"].tolist()
        me_names = [ link_name.split("-")[0] for link_name in link_names ]

        return me_names
    
    def is_correct_me(row: pd.Series, me_name: str) -> bool:
        link_name = row["Link Name"]
        return me_name in link_name
    
    def get_filtered_result(self, df_result: pd.DataFrame, selected_me_name: str) -> pd.DataFrame:
        if not selected_me_name:
            return df_result
        
        return df_result[df_result["Link Name"].str.contains(selected_me_name, na=False)]

    def process(self):
        if self.df_ref is not None and self.df_raw_data is not None:
            df_result = self.build_result_df()
            me_names = self.get_me_names(df_result)

            selected_me_name = st.selectbox(
                "Managed Element",
                me_names,
                index=None,
                placeholder="Choose options"
            )

            st.markdown(selected_me_name)

            df_filtered = self.get_filtered_result(df_result, selected_me_name)

            st.dataframe(df_filtered.style.apply(self.isDiffError, axis=1), hide_index=True)

            self.draw_color_legend()


# region Analyzer for Core
class CoreAnalyzer(EOLAnalyzer):
    def calculate_loss_between_core(self, df_result: pd.DataFrame) -> pd.DataFrame:
        forward_direction = df_result["Loss current - Loss EOL"].iloc[::2].values
        reverse_direction = df_result["Loss current - Loss EOL"].iloc[1::2].values
        loss_between_core = [abs(f - r) for f, r in zip(forward_direction, reverse_direction)]

        loss_between_core = [
            "--" if pd.isna(value) else round(value, 2)
            for value in loss_between_core
        ]

        df_loss_between_core = pd.DataFrame()
        df_loss_between_core["Loss between core"] = loss_between_core

        return df_loss_between_core
    
    @staticmethod
    def getColorCondition(value, threshold = 2) -> str:
        if value == "--":
            return "flapping"
        elif value > threshold:
            return "error"
        return ""

    def build_loss_table_body(self, link_names, loss_values) -> str:
        table_body = ""

        for i in range(len(link_names)):
            status = CoreAnalyzer.getColorCondition(loss_values[i // 2])
            color = LossAnalyzer.getColor(status)

            merged_cells = ""
            if i % 2 == 0:
                formated_value = loss_values[i // 2]
                if formated_value != "--":
                    formated_value = "{:.2f}".format(loss_values[i // 2])
                    
                merged_cells = f"""
                    <td style='border: 1px solid rgba(250,250,250,0.1); padding: 4px 8px; text-align: center; {color}' rowspan=2>
                        {formated_value}
                    </td>
                """.strip()

            table_body += f"""
                <tr>
                    <td style='border: 1px solid rgba(250,250,250,0.1); padding: 4px 8px; {color}'>
                        {link_names[i]}
                    </td>{merged_cells}
                </tr>
            """.strip()

        return table_body
    
    def build_loss_table(self, link_names, loss_values) -> str:
        table_body = self.build_loss_table_body(link_names, loss_values)

        html = f"""
            <div style="
                max-height: 500px; 
                overflow-y: auto; 
                border: 1px solid rgba(250, 250, 250, 0.1); 
                border-radius: 0.5rem;
                box-sizing: border-box;
            ">
                <table style="
                    border-collapse: collapse; 
                    width: 100%; 
                    text-align: left; 
                    font-family: 'Source Sans', sans-serif;
                    font-size: 14px;
                ">
                    <thead style="background-color: rgba(26,28,36,1); color: #fafafa;">
                        <tr>
                            <th style="border: 1px solid rgba(250,250,250,0.1); padding: 4px 8px;">Link Name</th>
                            <th style="border: 1px solid rgba(250,250,250,0.1); padding: 4px 8px;">Loss between core</th>
                        </tr>
                    </thead>
                    <tbody style="background-color: #0e1117; color: #fafafa;">
                        {table_body}
                    </tbody>
                </table>
            </div>
        """

        return html

    def process(self):
        if self.df_ref is not None and self.df_raw_data is not None:
            df_result = super().build_result_df()

            df_eol_ref: pd.DataFrame = super().extract_eol_ref(self.df_ref)
            df_loss_between_core = self.calculate_loss_between_core(df_result)

            link_names = df_eol_ref["Link Name"].tolist()
            loss_values = df_loss_between_core["Loss between core"].tolist()

            html = self.build_loss_table(link_names, loss_values)

            st.markdown(html, unsafe_allow_html=True)

            st.markdown("""
                <div style='display: flex; justify-content: center; align-items: center; gap: 16px; margin-bottom: 1rem'>
                    <div style='display: flex; justify-content: center; align-items: center; gap: 8px'>
                        <div style='background-color: #ff4d4d; width: 24px; height: 24px; border-radius: 8px;'></div>
                        <div style='text-align: center; color: #ff4d4d; font-size: 24px; font-weight: bold;'>
                            Loss not OK 
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
