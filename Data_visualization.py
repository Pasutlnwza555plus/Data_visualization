import streamlit as st 
import pandas as pd
import re
import plotly.express as px
from collections import defaultdict
# ตั้งชื่อไฟล์ชั่วคราว
UPLOAD_PATH_OPTICAL = "uploaded_optical.xlsx"
UPLOAD_PATH_FM = "uploaded_fm.xlsx"

pd.set_option("styler.render.max_elements", 1_200_000)
# สร้าง session state
if 'optical_uploaded' not in st.session_state:
    st.session_state.optical_uploaded = False
if 'fm_uploaded' not in st.session_state:
    st.session_state.fm_uploaded = False





#Line board




# Sidebar
menu = st.sidebar.radio("เลือกกิจกรรม", ["Loss between EOL", "หน้าแรก","CPU","FAN","MSU","Line board","Client board","Fiber Flapping","Loss between Core"])
if menu == "หน้าแรก":
    st.subheader("DWDM Monitoring Dashboard")
    
if menu == "CPU":
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

elif menu == "FAN":
    st.markdown("### 📂 Upload FAN File ")        
    uploaded_fan = st.file_uploader(" Upload FAN File", type=["xlsx"], key="fan")
    # ถ้าอัปโหลดใหม่ ให้โหลดและบันทึกใน session
    if uploaded_fan:
        df_fan = pd.read_excel(uploaded_fan)
        st.session_state.fan_data = df_fan  # เก็บลง session
        st.success("FAN file uploaded and stored")

    # ใช้ข้อมูลจาก session หากมี
    if st.session_state.get("fan_data") is not None:

        df_fan = st.session_state.fan_data.copy()
        try:
            # อ่านไฟล์ที่อัป
            df_fan = pd.read_excel("uploaded_fan.xlsx")
            df_fan.columns = df_fan.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.replace('\u00a0', ' ')
            # ตรวจสอบคอลัมน์ที่ต้องมี
            required_cols = {"ME", "Measure Object", "Begin Time", "End Time", "Value of Fan Rotate Speed(Rps)"}
            if not required_cols.issubset(df_fan.columns):
                st.error(f" ไฟล์ที่อัปโหลดต้องมีคอลัมน์: {', '.join(required_cols)}")
                st.write("คอลัมน์ที่พบ:", df_fan.columns.tolist())
                st.stop()
            #  Mapping Format 
            df_fan["Mapping Format"] = df_fan["ME"].astype(str).str.strip() + df_fan["Measure Object"].astype(str).str.strip()
            # โหลด reference FAN จากระบบ
            df_ref = pd.read_excel("data/FAN.xlsx")
            df_ref.columns = df_ref.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.replace('\u00a0', ' ')
            # เตรียมคอลัมน์สำคัญจาก df_ref
            df_ref_subset = df_ref[["Mapping", "Site Name", "Maximum threshold", "Minimun threshold"]].copy()
            df_ref_subset["Mapping"] = df_ref_subset["Mapping"].astype(str).str.strip()
            # รวมข้อมูลทั้งสอง DataFrame โดย match ด้วย Mapping Format
            df_merged = pd.merge(
                df_fan,
                df_ref_subset,
                left_on="Mapping Format",
                right_on="Mapping",
                how="inner"
            )
            # แสดงผลลัพธ์
            if not df_merged.empty:
                df_result = df_merged[[
                    "Begin Time", "End Time", "Site Name", "ME", "Measure Object",
                    "Maximum threshold", "Minimun threshold", "Value of Fan Rotate Speed(Rps)"
                ]]
                # 🔍 ตรวจสอบ FAN Performance โดยไม่ใส่ "Not OK" แต่เน้นสีแดง
                # สร้างคอลัมน์ flag ว่าแถวนี้เข้าเงื่อนไข Not OK หรือไม่
                #  สร้าง mask แยกสำหรับแถวที่ต้องไฮไลต์ โดยไม่ใส่ลงใน DataFrame
                def is_not_ok(row):
                    mo = str(row["Measure Object"])
                    val = row["Value of Fan Rotate Speed(Rps)"]
                    if "FCC" in mo and val > 120:
                        return True
                    elif "FCPP" in mo and val > 250:
                        return True
                    elif "FCPL" in mo and val > 120:
                        return True
                    elif "FCPS" in mo and val > 230:
                        return True
                    else:
                        return False
                # สร้าง boolean mask
                highlight_mask = df_result.apply(is_not_ok, axis=1)
                #  ฟังก์ชันไฮไลต์เฉพาะคอลัมน์เดียว
                def highlight_red_single_column(x):
                    return ['background-color: #ff4d4d; color: white' if highlight_mask.iloc[i] else '' for i in range(len(x))]
               
                styled_df = df_result.style.apply(
                    highlight_red_single_column,
                    subset=["Value of Fan Rotate Speed(Rps)"]
                )
                # ทศนิยม 2 ตำแหน่ง 
                styled_df = styled_df.format({
                    "Value of Fan Rotate Speed(Rps)": "{:.2f}"
                })

                st.markdown("### FAN Performance")
                st.dataframe(styled_df, use_container_width=True)
                if highlight_mask.any():
                    st.markdown("""
<div style='text-align: center; color: red; font-size: 24px; font-weight: bold;'>
FAN NOT OK
</div>
""", unsafe_allow_html=True)
                else:
                    st.markdown("""
<div style='text-align: center; color: green; font-size: 24px; font-weight: bold;'>
FAN OK
</div>
""", unsafe_allow_html=True)

            else:
                st.info("ไม่พบรายการที่ Mapping Format ตรงกับ FAN reference")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดระหว่างการประมวลผล: {e}")
    else:
        st.info("Please upload FAN ratio file")

elif menu == "MSU":
    st.markdown("### 📂 Upload MSU File")
    uploaded_msu = st.file_uploader("Upload MSU File", type=["xlsx"], key="msu")
    if uploaded_msu:
        df_msu = pd.read_excel(uploaded_msu)
        st.session_state.msu_data = df_msu
        st.success(" MSU file uploaded and stored")
    if st.session_state.get("msu_data") is not None:
        try:
            df_msu = st.session_state.msu_data.copy()
            df_msu.columns = df_msu.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.replace('\u00a0', ' ')
            # ตรวจสอบคอลัมน์ที่ต้องมี
            required_cols = {"ME", "Measure Object", "Laser Bias Current(mA)"}
            if not required_cols.issubset(df_msu.columns):
                st.error(f"❗ MSU file must contain columns: {', '.join(required_cols)}")
                st.stop()
            # สร้าง Mapping Format
            df_msu["Mapping Format"] = df_msu["ME"].astype(str).str.strip() + df_msu["Measure Object"].astype(str).str.strip()
            # โหลด Reference File
            df_ref = pd.read_excel("data/MSU.xlsx")
            df_ref.columns = df_ref.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.replace('\u00a0', ' ')
            # ตรวจสอบคอลัมน์ใน reference
            required_ref_cols = {"Mapping", "Maximum threshold"}
            if not required_ref_cols.issubset(df_ref.columns):
                st.error(f"❗ Reference file must contain columns: {', '.join(required_ref_cols)}")
                st.stop()
            df_ref["Mapping"] = df_ref["Mapping"].astype(str).str.strip()
            # Merge ข้อมูล
            df_merged = pd.merge(
                df_msu,
                df_ref[["Mapping", "Maximum threshold"]],
                left_on="Mapping Format",
                right_on="Mapping",
                how="inner"
            )
            if not df_merged.empty:
                df_result = df_merged[[
                    "ME", "Measure Object", "Maximum threshold", "Laser Bias Current(mA)"
                ]]
                # เช็คว่าเกิน threshold มั้ย
                def is_not_ok(row):
                    return row["Laser Bias Current(mA)"] > row["Maximum threshold"]
                highlight_mask = df_result.apply(is_not_ok, axis=1)
                # ฟังก์ชันเน้นสีแดง
                def highlight_red_single_column(x):
                    return ['background-color: #ff4d4d; color: white' if highlight_mask.iloc[i] else '' for i in range(len(x))]
                # แสดงผล
                styled_df = df_result.style.apply(
                    highlight_red_single_column,
                    subset=["Laser Bias Current(mA)"]
                ).format({
                    "Laser Bias Current(mA)": "{:.2f}",
                    "Maximum threshold": "{:.2f}"
                })
                st.markdown("### MSU Performance")
                st.dataframe(styled_df, use_container_width=True)
                # แสดงสถานะ
                if highlight_mask.any():
                    st.markdown("""
<div style='text-align: center; color: red; font-size: 24px; font-weight: bold;'>
🔴 MSU NOT OK
</div>
""", unsafe_allow_html=True)
                else:
                    st.markdown("""
<div style='text-align: center; color: green; font-size: 24px; font-weight: bold;'>
🟢 MSU OK
</div>
""", unsafe_allow_html=True)
            else:
                st.warning("ไม่พบ Mapping ที่ตรงกันระหว่าง MSU file และ reference")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดระหว่างการประมวลผล: {e}")
    else:
        st.info("กรุณาอัปโหลด MSU file ก่อนเพื่อเริ่มการวิเคราะห์")



elif menu == "Line board":
    st.markdown("### Upload Line cards performance File")
    uploaded_line = st.file_uploader("Upload Line cards File", type=["xlsx"], key="line")
    uploaded_log = st.file_uploader("Upload WASON Log", type=["txt"], key="log")
    def get_preset_map(log_text):
        import re
        lines = log_text.splitlines()
        ipmap = {
            "30.10.90.6": "HYI-4",
            "30.10.10.6": "Jasmine",
            "30.10.30.6": "Phu Nga",
            "30.10.50.6": "SNI-POI",
            "30.10.70.6": "NKS",
            "30.10.110.6": "PKT"
        }
        pmap = {}
        i = 0
        while i < len(lines):
            line = lines[i]
            m = re.search(r"\[CALL\s+\d+\]\s+\[([\d.]+)\s+[\d.]+\s+(\d+)\]", line)
            if m:
                ip = m.group(1).strip()
                cid = m.group(2).strip().lstrip("0")
                site = ipmap.get(ip, "Unknown")
                key = f"{cid} ({site})"
                j = i + 1
                while j < len(lines):
                    if "[CALL" in lines[j]:
                        break
                    if "[PreRout]:" in lines[j]:
                        k = j + 1
                        while k < len(lines):
                            if "[CALL" in lines[k]:
                                break
                            m2 = re.search(r"--(\d+)--WORK--\(USED\)--\(SUCCESS\)", lines[k])
                            if m2:
                                preset = m2.group(1).strip()
                                pmap[key] = preset
                                break
                            k += 1
                        break
                    j += 1
            i += 1
        return pmap
    if uploaded_log:
        log_text = uploaded_log.read().decode("utf-8", errors="ignore")
        pmap = get_preset_map(log_text)
    else:
        pmap = {}
    if uploaded_line:
        df_line = pd.read_excel(uploaded_line)
        st.session_state.line_data = df_line
        st.success("Line cards file uploaded and stored")
    if uploaded_log and st.session_state.get("line_data") is not None:
        try:
            df_line = st.session_state.line_data.copy()
            df_line.columns = df_line.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.replace('\u00a0', ' ')
            df_ref = pd.read_excel("data/Line.xlsx")
            df_ref.columns = df_ref.columns.astype(str)\
                .str.encode('ascii', 'ignore').str.decode('utf-8')\
                .str.replace(r'\s+', ' ', regex=True).str.strip()
            # ✅ เพิ่มลำดับไว้ใช้เรียงภายหลัง
            df_ref["order"] = range(len(df_ref))
            required_cols = {"ME", "Measure Object", "Instant BER After FEC", "Input Optical Power(dBm)", "Output Optical Power (dBm)"}
            if not required_cols.issubset(df_line.columns):
                st.error(f"Line cards file must contain columns: {', '.join(required_cols)}")
                st.stop()
            df_line["Mapping Format"] = df_line["ME"].astype(str).str.strip() + df_line["Measure Object"].astype(str).str.strip()
            df_ref["Mapping"] = df_ref["Mapping"].astype(str).str.strip()
            col_out = "Output Optical Power (dBm)"
            col_in = "Input Optical Power(dBm)"
            col_max_out = "Maximum threshold(out)"
            col_min_out = "Minimum threshold(out)"
            col_max_in = "Maximum threshold(in)"
            col_min_in = "Minimum threshold(in)"
            df_merged = pd.merge(
                df_line,
                df_ref[["Site Name", "Mapping", "Call ID", "Threshold", col_max_out, col_min_out, col_max_in, col_min_in, "Route", "order"]],
                left_on="Mapping Format",
                right_on="Mapping",
                how="inner"
            )
            if not df_merged.empty:
                df_result = df_merged[[ 
                    "Site Name", "ME", "Mapping", "Call ID", "Measure Object", "Threshold", "Instant BER After FEC",
                    col_max_out, col_min_out, col_out,
                    col_max_in, col_min_in, col_in, "Route", "order"
                ]]
                df_result["Call ID"] = df_result["Call ID"].astype(str).str.strip().str.lstrip("0")
                df_result["Route"] = df_result.apply(
                    lambda r: f"Preset {pmap[r['Call ID']]}" if r['Call ID'] in pmap else r["Route"],
                    axis=1
                )
                # ✅ เรียงตามลำดับจาก ref แล้วลบคอลัมน์ order ทิ้ง
                df_result = df_result.sort_values("order").drop(columns=["order"]).reset_index(drop=True)
               
                styled_df = df_result.style
                # ✅ ไฮไลต์ช่องที่เป็น Preset ด้วยสีน้ำเงิน
                styled_df = styled_df.apply(
                    lambda col: [
                        'background-color: lightblue; color: black'
                        if str(row["Route"]).startswith("Preset") else ''
                        for _, row in df_result.iterrows()
                    ],
                    subset=["Route"]
                )
                # ✅ ไฮไลต์ cell สีแดงเข้ม
                def highlight_critical_cells(val, colname, row):
                    if colname == "Instant BER After FEC":
                        try:
                            return 'background-color: #ff4d4d; color: white' if float(val) > 0 else ''
                        except:
                            return ''
                    elif colname == col_out:
                        try:
                            return 'background-color: #ff4d4d; color: white' if (val > row[col_max_out] or val < row[col_min_out]) else ''
                        except:
                            return ''
                    elif colname == col_in:
                        try:
                            return 'background-color: #ff4d4d; color: white' if (val > row[col_max_in] or val < row[col_min_in]) else ''
                        except:
                            return ''
                    return ''
                for colname in ["Instant BER After FEC", col_out, col_in]:
                    styled_df = styled_df.apply(
                        lambda col: [
                            highlight_critical_cells(row[colname], colname, row)
                            for _, row in df_result.iterrows()
                        ],
                        subset=[colname]
                    )
                # ✅ ไฮไลต์ cell สีแดงเข้ม
                def highlight_critical_cells(val, colname, row):
                    if colname == "Instant BER After FEC":
                        try:
                            return 'background-color: #ff4d4d; color: white' if float(val) > 0 else ''
                        except:
                            return ''
                    elif colname == col_out:
                        try:
                            return 'background-color: #ff4d4d; color: white' if (val > row[col_max_out] or val < row[col_min_out]) else ''
                        except:
                            return ''
                    elif colname == col_in:
                        try:
                            return 'background-color: #ff4d4d; color: white' if (val > row[col_max_in] or val < row[col_min_in]) else ''
                        except:
                            return ''
                    return ''
                # ✅ Apply cell-level red highlight (3 columns)
                for colname in ["Instant BER After FEC", col_out, col_in]:
                    styled_df = styled_df.apply(
                        lambda col: [
                            highlight_critical_cells(row[colname], colname, row)
                            for _, row in df_result.iterrows()
                        ],
                        subset=[colname]
                    )
                # ✅ ไฮไลต์แถวสีเทาอ่อน ถ้ามี cell critical
                def highlight_fail_rows(row):
                    fail = (
                        row["Instant BER After FEC"] > 0 or
                        row[col_out] > row[col_max_out] or row[col_out] < row[col_min_out] or
                        row[col_in] > row[col_max_in] or row[col_in] < row[col_min_in]
                    )
                    return ['background-color: #f2f2f2; color: black' if fail else '' for _ in row]
                styled_df = styled_df.apply(highlight_fail_rows, axis=1)
                # ✅ ไฮไลต์ Preset ด้วยสีน้ำเงิน
                styled_df = styled_df.apply(
                    lambda col: [
                        'background-color: lightblue; color: black'
                        if str(row["Route"]).startswith("Preset") else ''
                        for _, row in df_result.iterrows()
                    ],
                    subset=["Route"]
                )
                # ✅ จัดรูปแบบตัวเลข
                styled_df = styled_df.format({
                    col_out: "{:.2f}",
                    col_in: "{:.2f}",
                    "Instant BER After FEC": "{:.1e}"
                })
                st.markdown("### Line Performance")
                st.dataframe(styled_df, use_container_width=True)
                # ✅ ตรวจแถวที่ผิดเพื่อแสดงผล OK / NOT OK
                failed_rows = df_result.apply(
                    lambda row: (
                        row[col_out] > row[col_max_out] or row[col_out] < row[col_min_out] or
                        row[col_in] > row[col_max_in] or row[col_in] < row[col_min_in] or
                        row["Instant BER After FEC"] > 0
                    ),
                    axis=1
                )
                if failed_rows.any():
                    st.markdown("<div style='text-align: center; color: red; font-size: 24px; font-weight: bold;'>Line NOT OK</div>", unsafe_allow_html=True)
                else:
                    st.markdown("<div style='text-align: center; color: green; font-size: 24px; font-weight: bold;'>Line OK</div>", unsafe_allow_html=True)
            else:
                st.warning("ไม่พบ Mapping ที่ตรงกันระหว่าง Line file และ reference")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดระหว่างการประมวลผล: {e}")
    else:
        st.info("กรุณาอัปโหลด Line file ก่อน")






elif menu == "Fiber Flapping":
    st.markdown("### Upload OSC & FM Files")
    # Upload OSC
    uploaded_optical = st.file_uploader("Upload OSC Optical File", type=["xlsx"], key="osc")
    if uploaded_optical:
        df_optical = pd.read_excel(uploaded_optical)
        st.session_state.osc_optical_data = df_optical
        st.success("OSC Optical File Uploaded")
    # Upload FM
    uploaded_fm = st.file_uploader("Upload FM Alarm File", type=["xlsx"], key="fm")
    if uploaded_fm:
        df_fm = pd.read_excel(uploaded_fm)
        st.session_state.osc_fm_data = df_fm
        st.success("FM Alarm File Uploaded")
    # ประมวลผลเมื่อทั้งสองไฟล์ถูกอัปโหลดแล้ว
    if "osc_optical_data" in st.session_state and "osc_fm_data" in st.session_state:
        try:
            df_optical = st.session_state.osc_optical_data.copy()
            df_fm = st.session_state.osc_fm_data.copy()
            df_optical.columns = df_optical.columns.str.strip()
            df_optical["Max - Min (dB)"] = df_optical["Max Value of Input Optical Power(dBm)"] - df_optical["Min Value of Input Optical Power(dBm)"]
            # Extract Target ME
            def extract_target(measure_obj):
                match = re.search(r'\(([^)]+)\)', str(measure_obj))
                return match.group(1) if match else None
            df_optical["Target ME"] = df_optical["Measure Object"].apply(extract_target)
            df_optical["Begin Time"] = pd.to_datetime(df_optical["Begin Time"], errors="coerce")
            df_optical["End Time"] = pd.to_datetime(df_optical["End Time"], errors="coerce")
            # Filter > 2dB
            df_filtered = df_optical[df_optical["Max - Min (dB)"] > 2].copy()
            df_fm.columns = df_fm.columns.str.strip()
            df_fm["Occurrence Time"] = pd.to_datetime(df_fm["Occurrence Time"], errors="coerce")
            df_fm["Clear Time"] = pd.to_datetime(df_fm["Clear Time"], errors="coerce")
            link_col = [col for col in df_fm.columns if col.startswith("Link")][0]
            # หา no-match
            result = []
            for _, row in df_filtered.iterrows():
                me = re.escape(str(row["ME"]))
                target_me = re.escape(str(row["Target ME"]))
                matched = df_fm[
                    df_fm[link_col].astype(str).str.contains(me, na=False) &
                    df_fm[link_col].astype(str).str.contains(target_me, na=False) &
                    (df_fm["Occurrence Time"] <= row["End Time"]) &
                    (df_fm["Clear Time"] >= row["Begin Time"])
                ]
                if matched.empty:
                    result.append(row)
            df_nomatch = pd.DataFrame(result)
            # Show result
            st.markdown("### OSC Power Flapping (No Alarm Match)")
            if not df_nomatch.empty:
                st.dataframe(
                    df_nomatch[[
                        "Begin Time", "End Time", "Granularity", "ME", "ME IP", "Measure Object",
                        "Max Value of Input Optical Power(dBm)", "Min Value of Input Optical Power(dBm)",
                        "Input Optical Power(dBm)", "Max - Min (dB)"
                    ]],
                    use_container_width=True
                )
                # Daily graph
                df_nomatch["Date"] = df_nomatch["Begin Time"].dt.date
                site_count = df_nomatch.groupby("Date")["ME"].nunique().reset_index()
                site_count.columns = ["Date", "Sites"]
                import plotly.express as px
                fig = px.bar(site_count, x="Date", y="Sites", text="Sites", title="No Fiber Break Alarm Match")
                fig.update_traces(marker_color="crimson", textposition="outside")
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
                # แสดงแยกรายวัน
                for date, group in df_nomatch.groupby("Date"):
                    st.markdown(f"#### {date.strftime('%Y-%m-%d')}")
                    st.dataframe(group[[
                        "Begin Time", "End Time", "Granularity", "ME", "ME IP", "Measure Object",
                        "Max Value of Input Optical Power(dBm)", "Min Value of Input Optical Power(dBm)",
                        "Input Optical Power(dBm)", "Max - Min (dB)"
                    ]], use_container_width=True)
            else:
                st.success("ไม่มีรายการ Flapping ที่ไม่แมตช์ FM")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
    else:
        st.info("กรุณาอัปโหลดทั้ง OSC และ FM ไฟล์ก่อน")



# Loss in Core
elif menu == "Loss between EOL":
    st.markdown("### Please upload files")

    EOL_sheet_name = "Loss between core & EOL"

    def get_df_recent_rank(df_ref:pd.DataFrame, recent_rank: int = 0) -> pd.DataFrame:
        header_len = len(df_ref.columns)
        if (header_len - 4*recent_rank < 12):
            raise Exception("Data not found")

        start = -4 - 4*recent_rank
        end_col = header_len if recent_rank == 0 else -4 * recent_rank
        header_names = df_ref.columns[start:end_col].to_list()
        eol_ref_columns = df_ref.columns[df_ref.iloc[0] == "EOL(dB)"]

        df_date_ref = pd.to_numeric(df_ref[header_names[0]], downcast="float", errors="coerce")
        df_eol_ref = pd.to_numeric(df_ref[eol_ref_columns[0]], downcast="float", errors="coerce")

        calculated_diff = df_date_ref - df_eol_ref - 1

        df_eol = pd.DataFrame()

        df_eol["Link Name"] = df_ref['140.1'].iloc[1:]
        df_eol["EOL(dB)"] = df_eol_ref
        df_eol["Current Attenuation(dB)"] = df_ref[header_names[0]]
        df_eol["Loss current - Loss EOL"] = calculated_diff
        df_eol["Remark"] = df_ref[header_names[3]]

        return df_eol
        
    def isDiffError(row):
        return ['background-color: #ff4d4d; color: white'] * len(row) if float(row["Loss current - Loss EOL"]) >= 2 else [''] * len(row)


    uploaded_reference = st.file_uploader("Upload Data Sheet", type=["xlsx"], key="ref")
    if uploaded_reference:
        df_ref = pd.read_excel(uploaded_reference, sheet_name=EOL_sheet_name)

        
        # col_names_primary   = df_ref.iloc[0].to_list()
        # col_names_secondary = df_ref.iloc[1].to_list()
        # try:
        df_eol = get_df_recent_rank(df_ref, 1)

        st.dataframe(df_eol.style.apply(isDiffError, axis=1), hide_index=True)
        
        # except Exception: 
        #     st.markdown("Data not found")

        # st.markdown(header_names[-4:])
        # st.markdown(col_names_primary[-4:])
        # st.markdown(col_names_secondary[-4:])

        # st.markdown()

        # st.session_state.reference_sheet = df_ref
        # st.success("OSC Optical File Uploaded")
    