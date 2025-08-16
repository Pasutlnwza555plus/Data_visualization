import streamlit as st
import pandas as pd

def cpu_page():
    st.markdown("### Upload CPU File")
    # Upload CPU File
    uploaded_cpu = st.file_uploader("Upload CPU File", type=["xlsx"], key="cpu")
    if uploaded_cpu:
        df_cpu = pd.read_excel(uploaded_cpu)
        st.session_state.cpu_data = df_cpu
        st.success("CPU file uploaded and stored")
    # ใช้จาก session ถ้ามีข้อมูล
    if st.session_state.get("cpu_data") is not None:
        try:
            # เตรียม DataFrame
            df_cpu = st.session_state.cpu_data.copy()
            df_cpu.columns = df_cpu.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.replace('\u00a0', ' ')
            # ตรวจสอบคอลัมน์ที่จำเป็น
            required_cols = {"ME", "Measure Object", "CPU utilization ratio"}
            if not required_cols.issubset(df_cpu.columns):
                st.error(f"CPU file must contain columns: {', '.join(required_cols)}")
                st.stop()
            # สร้าง Mapping Format
            df_cpu["Mapping Format"] = df_cpu["ME"].astype(str).str.strip() + df_cpu["Measure Object"].astype(str).str.strip()
            # โหลด Reference File
            df_ref = pd.read_excel("data/CPU.xlsx")
            df_ref.columns = df_ref.columns.astype(str)\
                            .str.encode('ascii', 'ignore').str.decode('utf-8')\
                                .str.replace(r'\s+', ' ', regex=True)\
                                .str.strip()
            # ตรวจสอบคอลัมน์ใน reference
            required_ref_cols = {"Mapping", "Maximum threshold", "Minimum threshold"}
            if not required_ref_cols.issubset(df_ref.columns):
                st.error(f"Reference file must contain columns: {', '.join(required_ref_cols)}")
                st.stop()
            df_ref["Mapping"] = df_ref["Mapping"].astype(str).str.strip()
            # Merge ข้อมูล
            df_merged = pd.merge(
                df_cpu,
                df_ref[["Site Name", "Mapping", "Maximum threshold", "Minimum threshold"]],
                left_on="Mapping Format",
                right_on="Mapping",
                how="inner"
            )
            # ตรวจสอบผลลัพธ์
            if not df_merged.empty:
                df_result = df_merged[[
                    "Site Name", "ME", "Measure Object", "Maximum threshold", "Minimum threshold","CPU utilization ratio"
                ]]
                # ฟังก์ชันตรวจไม่ผ่าน
                def is_not_ok(row):
                    return row["CPU utilization ratio"] > row["Maximum threshold"] or row["CPU utilization ratio"] < row["Minimum threshold"]
                highlight_mask = df_result.apply(is_not_ok, axis=1)
                # ฟังก์ชันไฮไลต์สีแดงเฉพาะ column
                def highlight_red_single_column(x):
                    return ['background-color: #ff4d4d; color: white' if highlight_mask.iloc[i] else '' for i in range(len(x))]
                # สร้าง styled table
                styled_df = df_result.style.apply(
                    highlight_red_single_column,
                    subset=["CPU utilization ratio"]
                ).format({
                    "CPU utilization ratio": "{:.2f}",
                    "Maximum threshold": "{:.2f}",
                    "Minimum threshold": "{:.2f}"
                })
                st.markdown("### CPU Performance")
                st.dataframe(styled_df, use_container_width=True)
                # แสดงผลสรุป
                if highlight_mask.any():
                    st.markdown("""
                        <div style='text-align: center; color: red; font-size: 24px; font-weight: bold;'>
                            CPU NOT OK
                        </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                        <div style='text-align: center; color: green; font-size: 24px; font-weight: bold;'>
                            CPU OK
                        </div>
                    """, unsafe_allow_html=True)
            else:
                st.warning("ไม่พบ Mapping ที่ตรงกันระหว่าง CPU file และ reference")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดระหว่างการประมวลผล: {e}")
    else:
        st.info("กรุณาอัปโหลด CPU file ก่อนเพื่อเริ่มการวิเคราะห์")