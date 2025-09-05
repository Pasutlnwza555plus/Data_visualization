import streamlit as st 
import pandas as pd
import re
import plotly.express as px
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional

from pages.loss import EOLAnalyzer, CoreAnalyzer
from services.database import Database

st.set_page_config(layout="wide")

pd.set_option("styler.render.max_elements", 1_200_000)

database_service = Database()

st.session_state.reference_sheet = database_service.get_reference_sheet()

#filter
def cascading_filter(
    df: pd.DataFrame,
    cols: list[str],
    *,
    ns: str = "flt",
    labels: dict[str, str] | None = None,
    clear_text: str = "Clear Filters",
):
   
    if labels is None:
        labels = {}

    # เตรียม state
    for c in cols:
        st.session_state.setdefault(f"{ns}_f_{c}", [])

    # คอลัมน์ที่มีอยู่จริงเท่านั้น
    active_cols = [c for c in cols if c in df.columns]
    if not active_cols:
        return df.reset_index(drop=True), {}

    # สร้าง options ทีละชั้นด้วย mask สะสม
    masks = [pd.Series(True, index=df.index)]
    options_per_col = []
    for i, c in enumerate(active_cols):
        m = masks[-1]
        opts = sorted(df.loc[m, c].dropna().astype(str).unique())
        options_per_col.append(opts)

        # prune ค่าเลือกที่ไม่อยู่ใน opts (กัน selection ค้าง)
        valid_sel = [x for x in st.session_state[f"{ns}_f_{c}"] if x in opts]
        st.session_state[f"{ns}_f_{c}"] = valid_sel

        # อัปเดต mask สำหรับคอลัมน์ถัดไป
        if valid_sel:
            masks.append(m & df[c].astype(str).isin(valid_sel))
        else:
            masks.append(m)

    # วาด widgets เป็นแถวเดียว + ปุ่ม Clear
    cols_widgets = st.columns([1] * len(active_cols) + [0.8])
    for i, c in enumerate(active_cols):
        with cols_widgets[i]:
            st.multiselect(
                labels.get(c, c),
                options_per_col[i],
                key=f"{ns}_f_{c}",
            )

    def _clear():
        for c in active_cols:
            st.session_state.pop(f"{ns}_f_{c}", None)

    with cols_widgets[-1]:
        st.button(clear_text, on_click=_clear)

    # สร้าง final mask จาก selections ทั้งหมด
    final_mask = pd.Series(True, index=df.index)
    selections = {}
    for c in active_cols:
        sel = st.session_state.get(f"{ns}_f_{c}", [])
        selections[c] = sel
        if sel:
            final_mask &= df[c].astype(str).isin(sel)

    return df[final_mask].reset_index(drop=True), selections

#Preset status
# Match: [WASON][CALL 8] [30.10.90.6 30.10.10.6 85] COPPER
CALL_HEADER_RE = re.compile(r"\[WASON\]\[CALL\s+(\d+)\]\s+\[([^\]]+)\]")
# Any Conn line that contains WR
CONN_HAS_WR_RE = re.compile(r"\[WASON\]\s*\[Conn\s+\d+\].*\bWR\b", re.IGNORECASE)
# Specifically "WR NO_ALARM"
CONN_WR_NOALARM_RE = re.compile(r"\[WASON\]\s*\[Conn\s+\d+\][^\n]*\bWR\s+NO_ALARM\b", re.IGNORECASE)
# A preroute line like: --2--WORK--(USED)--(SUCCESS)-- (be robust to trailing text)
PREROUT_USED_RE = re.compile(
    r"\[WASON\]--\s*(\d+)\s*--\s*WORK\s*--\s*\(USED\)\s*--\s*\((\w+)\).*",
    re.IGNORECASE,
)

@dataclass
class CallBlock:
    call_id: int
    triple: str
    lines: List[str] = field(default_factory=list)

def parse_calls(text: str) -> List[CallBlock]:
    calls: List[CallBlock] = []
    cur: Optional[CallBlock] = None
    for raw in text.splitlines():
        line = raw.rstrip("\n")
        m = CALL_HEADER_RE.search(line)
        if m:
            if cur is not None:
                calls.append(cur)
            cur = CallBlock(call_id=int(m.group(1)), triple=m.group(2), lines=[])
        if cur is not None:
            cur.lines.append(line)
    if cur is not None:
        calls.append(cur)
    return calls

def evaluate_preset_status(cb: CallBlock) -> Dict[str, Any]:
    """
    Focused rules:
    - Consider only CALLs that have WR (any Conn line includes 'WR').
    - Require at least one 'WR NO_ALARM' Conn line.
    - In [PreRout], there must be exactly one 'WORK (USED) (SUCCESS)' line.
    """
    has_wr = any(CONN_HAS_WR_RE.search(ln) for ln in cb.lines)
    if not has_wr:
        return {"has_wr": False}

    wr_no_alarm = any(CONN_WR_NOALARM_RE.search(ln) for ln in cb.lines)

    used_rows = []
    for ln in cb.lines:
        m = PREROUT_USED_RE.search(ln)
        if m:
            used_rows.append({"index": int(m.group(1)), "result": m.group(2).upper(), "raw": ln})

    verdict = "FAIL"
    Restore = ""
    pr_index: Optional[int] = None

    if not wr_no_alarm:
        Restore = "WR found but not WR NO_ALARM"
    elif len(used_rows) != 1:
       Restore = f"Found {len(used_rows)} USED rows (expected 1)"
    elif used_rows[0]["result"] != "SUCCESS":
        Restore = "USED row is not SUCCESS"
    else:
        verdict = "PASS"
        pr_index = used_rows[0]["index"]
        Restore = "Normal"

    return {
        "has_wr": True,
        "wr_no_alarm": wr_no_alarm,
        "verdict": verdict,
        "Restore": Restore,
        "pr_index": pr_index,
        "used_rows": used_rows,
        "raw": "\n".join(cb.lines),
    }

# Sidebar
menu = st.sidebar.radio("เลือกกิจกรรม", ["หน้าแรก","CPU","FAN","MSU","Line board","Client board","Fiber Flapping","Loss between Core","Loss between EOL","Preset status","Reference Sheet"])

if menu == "หน้าแรก":
    st.subheader("DWDM Monitoring Dashboard")


if menu == "CPU":
    st.markdown("### Upload CPU File")

    # Upload & cache
    uploaded_cpu = st.file_uploader("Upload CPU File", type=["xlsx"], key="cpu")
    if uploaded_cpu:
        st.session_state.cpu_data = pd.read_excel(uploaded_cpu)
        st.success("CPU file uploaded and stored")

    # Process
    if st.session_state.get("cpu_data") is not None:
        try:
            df_cpu = st.session_state.cpu_data.copy()
            # ทำความสะอาดชื่อคอลัมน์ (ให้แนวทางเดียวกับ line)
            df_cpu.columns = (
                df_cpu.columns.astype(str)
                .str.strip().str.replace(r"\s+", " ", regex=True)
                .str.replace("\u00a0", " ")
            )

            # ตรวจคอลัมน์ที่จำเป็น
            required_cols = {"ME", "Measure Object", "CPU utilization ratio"}
            if not required_cols.issubset(df_cpu.columns):
                st.error(f"CPU file must contain columns: {', '.join(sorted(required_cols))}")
                st.stop()

            # สร้าง Mapping Format
            df_cpu["Mapping Format"] = (
                df_cpu["ME"].astype(str).str.strip()
                + df_cpu["Measure Object"].astype(str).str.strip()
            )

            # โหลด reference
            df_ref = pd.read_excel("data/CPU.xlsx")
            df_ref.columns = (
                df_ref.columns.astype(str)
                .str.encode("ascii", "ignore").str.decode("utf-8")
                .str.replace(r"\s+", " ", regex=True)
                .str.strip()
            )

            required_ref_cols = {"Mapping", "Maximum threshold", "Minimum threshold"}
            if not required_ref_cols.issubset(df_ref.columns):
                st.error(f"Reference file must contain columns: {', '.join(sorted(required_ref_cols))}")
                st.stop()

            df_ref["Mapping"] = df_ref["Mapping"].astype(str).str.strip()
            # ✅ สร้าง order ตามลำดับในไฟล์ ref (เหมือน line)
            df_ref["order"] = range(len(df_ref))

            # เลือกคอลัมน์จาก ref ที่จะ merge (เผื่อมี Site/CallID/Route)
            ref_cols = ["Mapping", "Maximum threshold", "Minimum threshold", "order"]
            for extra in ["Site Name", "Call ID", "Route"]:
                if extra in df_ref.columns:
                    ref_cols.append(extra)

            # Merge
            df_merged = pd.merge(
                df_cpu,
                df_ref[ref_cols],
                left_on="Mapping Format",
                right_on="Mapping",
                how="inner"
            )

            if df_merged.empty:
                st.warning("No matching mapping found between CPU file and reference")
            else:
                base_cols = [
                    "ME", "Measure Object",
                    "Maximum threshold", "Minimum threshold",
                    "CPU utilization ratio", "order"
                ]
                opt_cols = [c for c in ["Site Name", "Call ID", "Route"] if c in df_merged.columns]
                show_cols = opt_cols + base_cols

                df_result = df_merged[show_cols].copy()
                # ✅ เรียงตาม order จาก ref แล้วลบทิ้ง
                df_result = df_result.sort_values("order").drop(columns=["order"]).reset_index(drop=True)

                # ✅ ใช้ฟิลเตอร์แบบไล่ชั้น (ไม่เหมือน line board)
                df_filtered, _sel = cascading_filter(
                    df_result,
                    cols=["Site Name", "ME", "Measure Object"],
                    ns="cpu",
                    clear_text="Clear CPU Filters"
                )
                st.caption(f"CPU (showing {len(df_filtered)}/{len(df_result)} rows)")

                # ตัวเลขสำหรับเปรียบเทียบ/format
                for c in ["CPU utilization ratio", "Maximum threshold", "Minimum threshold"]:
                    if c in df_filtered.columns:
                        df_filtered[c] = pd.to_numeric(df_filtered[c], errors="coerce")

                df_view = df_filtered.copy()

                # เงื่อนไข “ทั้งแถวเป็นเทา”
                def row_has_issue(r):
                    v  = r.get("CPU utilization ratio")
                    lo = r.get("Minimum threshold")
                    hi = r.get("Maximum threshold")
                    try:
                        return (pd.notna(v) and pd.notna(lo) and pd.notna(hi) and (v < lo or v > hi))
                    except:
                        return False

                styled = (
                    df_view.style
                    # เทาทั้งแถวถ้ามีปัญหา
                    .apply(lambda r: ['background-color:#e6e6e6;color:black' if row_has_issue(r) else '' for _ in r], axis=1)
                    # แดงเฉพาะช่อง CPU utilization ratio ที่นอก threshold
                    .apply(lambda _:
                        ['background-color:#ff4d4d;color:white'
                         if (pd.notna(v) and pd.notna(hi) and pd.notna(lo) and (v > hi or v < lo))
                         else ''
                         for v, hi, lo in zip(
                             df_view.get("CPU utilization ratio", pd.Series(index=df_view.index)),
                             df_view.get("Maximum threshold", pd.Series(index=df_view.index)),
                             df_view.get("Minimum threshold", pd.Series(index=df_view.index)),
                         )
                    ], subset=["CPU utilization ratio"] if "CPU utilization ratio" in df_view.columns else [])
                    # ฟ้าให้ Route ที่เป็น Preset (ถ้ามี)
                    .apply(lambda _:
                        ['background-color:lightblue;color:black' if str(x).startswith("Preset") else ''
                         for x in df_view["Route"]] if "Route" in df_view.columns else [],
                        subset=["Route"] if "Route" in df_view.columns else []
                    )
                    .format({
                        "CPU utilization ratio": "{:.2f}",
                        "Maximum threshold": "{:.2f}",
                        "Minimum threshold": "{:.2f}",
                    })
                )

                st.markdown("### CPU Performance")
                st.write(styled)

                # สรุปสถานะ (อิงข้อมูลหลังฟิลเตอร์)
                failed_rows = df_view.apply(row_has_issue, axis=1)
                st.markdown(
                    "<div style='text-align:center; font-size:32px; font-weight:bold; color:{};'>CPU Performance {}</div>".format(
                        "red" if failed_rows.any() else "green",
                        "Warning" if failed_rows.any() else "Normal"
                    ),
                    unsafe_allow_html=True
                )

        except Exception as e:
            st.error(f"An error occurred during processing: {e}")
    else:
        st.info("Please upload file to start the analysis")


elif menu == "FAN":
    st.markdown("### Upload FAN File")
    uploaded_fan = st.file_uploader("Upload FAN File", type=["xlsx"], key="fan")

    # cache to session
    if uploaded_fan:
        st.session_state.fan_data = pd.read_excel(uploaded_fan)
        st.success("FAN file uploaded and stored")

    # use from session
    if st.session_state.get("fan_data") is not None:
        try:
            # ---------- Prepare data ----------
            df_fan = st.session_state.fan_data.copy()
            df_fan.columns = (
                df_fan.columns.astype(str)
                .str.strip().str.replace(r"\s+", " ", regex=True)
                .str.replace("\u00a0", " ")
            )

            required_cols = {"ME", "Measure Object", "Begin Time", "End Time", "Value of Fan Rotate Speed(Rps)"}
            if not required_cols.issubset(df_fan.columns):
                st.error(f"Uploaded file must contain columns: {', '.join(sorted(required_cols))}")
                st.write("Detected columns:", df_fan.columns.tolist())
                st.stop()

            # Mapping key (unchanged)
            df_fan["Mapping Format"] = (
                df_fan["ME"].astype(str).str.strip()
                + df_fan["Measure Object"].astype(str).str.strip()
            )

            # ---------- Reference ----------
            df_ref = pd.read_excel("data/FAN.xlsx")
            df_ref.columns = (
                df_ref.columns.astype(str)
                .str.strip().str.replace(r"\s+", " ", regex=True)
                .str.replace("\u00a0", " ")
            )

            # keep important columns; note: spelling 'Minimun threshold' preserved per your file
            df_ref_subset = df_ref[["Mapping", "Site Name", "Maximum threshold", "Minimun threshold"]].copy()
            df_ref_subset["Mapping"] = df_ref_subset["Mapping"].astype(str).str.strip()

            # ✅ add reference order (like Line/CPU) for display sorting
            df_ref_subset["order"] = range(len(df_ref_subset))

            # ---------- Merge ----------
            df_merged = pd.merge(
                df_fan,
                df_ref_subset,
                left_on="Mapping Format",
                right_on="Mapping",
                how="inner"
            )

            if not df_merged.empty:
                df_result = df_merged[[
                    "Begin Time", "End Time", "Site Name", "ME", "Measure Object",
                    "Maximum threshold", "Minimun threshold", "Value of Fan Rotate Speed(Rps)", "order"
                ]].copy()

                # ✅ sort by reference order for consistent display
                df_result = df_result.sort_values("order").drop(columns=["order"]).reset_index(drop=True)

                # ---------- Cascading filters (like CPU) ----------
                df_filtered, _sel = cascading_filter(
                    df_result,
                    cols=["Site Name", "ME", "Measure Object"],
                    ns="fan",
                    clear_text="Clear FAN Filters"
                )
                st.caption(f"FAN (showing {len(df_filtered)}/{len(df_result)} rows)")

                # ---------- Styling ----------
                # numeric cast for display
                if "Value of Fan Rotate Speed(Rps)" in df_filtered.columns:
                    df_filtered["Value of Fan Rotate Speed(Rps)"] = pd.to_numeric(
                        df_filtered["Value of Fan Rotate Speed(Rps)"], errors="coerce"
                    )

                def is_not_ok(row):
                    """KEEP ORIGINAL LOGIC: FCC/FCPP/FCPL/FCPS thresholds"""
                    mo = str(row["Measure Object"])
                    val = row["Value of Fan Rotate Speed(Rps)"]
                    if pd.isna(val):
                        return False
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

                df_view = df_filtered.copy()
                highlight_mask = df_view.apply(is_not_ok, axis=1)

                # gray whole row when issue
                def gray_row(r):
                    return ['background-color:#e6e6e6;color:black' if highlight_mask.iloc[r.name] else '' for _ in r]

                # red only the value column when issue
                def red_value(_):
                    return ['background-color:#ff4d4d;color:white' if m else '' for m in highlight_mask]

                styled_df = (
                    df_view.style
                    .apply(gray_row, axis=1)
                    .apply(red_value, subset=["Value of Fan Rotate Speed(Rps)"])
                    .format({"Value of Fan Rotate Speed(Rps)": "{:.2f}"})
                )

                st.markdown("### FAN Performance")
                st.write(styled_df, use_container_width=True)

                # ---------- Status banner (consistent wording) ----------
                st.markdown(
                    "<div style='text-align:center; font-size:32px; font-weight:bold; color:{};'>FAN Performance {}</div>".format(
                        "red" if highlight_mask.any() else "green",
                        "Warning" if highlight_mask.any() else "Normal"
                    ),
                    unsafe_allow_html=True
                )

            else:
                st.info("No matching mapping found between FAN file and reference")

        except Exception as e:
            st.error(f"An error occurred during processing: {e}")
    else:
        st.info("Please upload a FAN file to start the analysis")


elif menu == "MSU":
    st.markdown("### Upload MSU File")

    uploaded_msu = st.file_uploader("Upload MSU File", type=["xlsx"], key="msu")
    if uploaded_msu:
        df_msu = pd.read_excel(uploaded_msu)
        st.session_state.msu_data = df_msu
        st.success("MSU file uploaded and stored")

    if st.session_state.get("msu_data") is not None:
        try:
            df_msu = st.session_state.msu_data.copy()
            df_msu.columns = (
                df_msu.columns.astype(str)
                .str.strip().str.replace(r'\s+', ' ', regex=True)
                .str.replace('\u00a0', ' ')
            )

            # ตรวจสอบคอลัมน์ที่ต้องมี
            required_cols = {"ME", "Measure Object", "Laser Bias Current(mA)"}
            if not required_cols.issubset(df_msu.columns):
                st.error(f"MSU file must contain columns: {', '.join(sorted(required_cols))}")
                st.stop()

            # สร้าง Mapping key
            df_msu["Mapping Format"] = (
                df_msu["ME"].astype(str).str.strip() +
                df_msu["Measure Object"].astype(str).str.strip()
            )

            # โหลด Reference
            df_ref = pd.read_excel("data/MSU.xlsx")
            df_ref.columns = (
                df_ref.columns.astype(str)
                .str.strip().str.replace(r'\s+', ' ', regex=True)
                .str.replace('\u00a0', ' ')
            )

            # ตรวจสอบคอลัมน์ใน reference
            required_ref_cols = {"Site Name", "Mapping", "Maximum threshold"}
            if not required_ref_cols.issubset(df_ref.columns):
                st.error(f"Reference file must contain columns: {', '.join(sorted(required_ref_cols))}")
                st.stop()

            df_ref["Mapping"] = df_ref["Mapping"].astype(str).str.strip()
            df_ref["order"] = range(len(df_ref))

            # Merge
            df_merged = pd.merge(
                df_msu,
                df_ref[["Site Name", "Mapping", "Maximum threshold", "order"]],
                left_on="Mapping Format",
                right_on="Mapping",
                how="inner"
            )

            if not df_merged.empty:
                # เลือกคอลัมน์ที่ต้องใช้ + เรียงตาม order
                df_result = (
                    df_merged[["Site Name", "ME", "Measure Object", "Maximum threshold", "Laser Bias Current(mA)", "order"]]
                    .sort_values("order")
                    .drop(columns=["order"])
                    .reset_index(drop=True)
                )

                # Filter
                df_filtered, _sel = cascading_filter(
                    df_result,
                    cols=["Site Name", "ME", "Measure Object"],
                    ns="msu",
                    clear_text="Clear MSU Filters"
                )
                st.caption(f"MSU (showing {len(df_filtered)}/{len(df_result)} rows)")

                # แปลงเป็นตัวเลข
                for c in ["Laser Bias Current(mA)", "Maximum threshold"]:
                    df_filtered[c] = pd.to_numeric(df_filtered[c], errors="coerce")

                # ฟังก์ชันตรวจเงื่อนไข
                def is_not_ok(row):
                    return row["Laser Bias Current(mA)"] > row["Maximum threshold"]

                # ไฮไลท์: แดงเฉพาะคอลัมน์ Laser Bias ที่ผิด
                def red_value(_):
                    return [
                        'background-color:#ff4d4d;color:white' if is_not_ok(row) else ''
                        for _, row in df_filtered.iterrows()
                    ]

                styled_df = (
                    df_filtered.style
                    .apply(red_value, subset=["Laser Bias Current(mA)"])
                    .format({
                        "Laser Bias Current(mA)": "{:.2f}",
                        "Maximum threshold": "{:.2f}",
                    })
                )

                st.markdown("### MSU Performance")
                st.write(styled_df, use_container_width=True)

                # Status
                failed_rows = df_filtered.apply(is_not_ok, axis=1)
                st.markdown(
                    "<div style='text-align:center; font-size:32px; font-weight:bold; color:{};'>MSU Performance {}</div>".format(
                        "red" if failed_rows.any() else "green",
                        "Warning" if failed_rows.any() else "Normal"
                    ),
                    unsafe_allow_html=True
                )

            else:
                st.warning("No matching mapping found between MSU file and reference")

        except Exception as e:
            st.error(f"An error occurred during processing: {e}")
    else:
        st.info("Please upload an MSU file to start the analysis")


elif menu == "Client board":
    st.markdown("### Upload Client File")

    # Upload Client File
    uploaded_client = st.file_uploader("Upload Client File", type=["xlsx"], key="client")
    if uploaded_client:
        df_client = pd.read_excel(uploaded_client)
        st.session_state.client_data = df_client
        st.success("Client file uploaded and stored")

    # ใช้จาก session ถ้ามีข้อมูล
    if st.session_state.get("client_data") is not None:
        try:
            # เตรียม DataFrame
            df_client = st.session_state.client_data.copy()
            df_client.columns = (
                df_client.columns.astype(str)
                .str.strip().str.replace(r'\s+', ' ', regex=True)
                .str.replace('\u00a0', ' ')
            )

            # ตรวจสอบคอลัมน์ที่จำเป็น
            required_cols = {"ME", "Measure Object", "Input Optical Power(dBm)", "Output Optical Power (dBm)"}
            if not required_cols.issubset(df_client.columns):
                st.error(f"Client file must contain columns: {', '.join(required_cols)}")
                st.stop()

            # สร้าง Mapping Format (คงตรรกะเดิม)
            df_client["Mapping Format"] = (
                df_client["ME"].astype(str).str.strip()
                + df_client["Measure Object"].astype(str).str.strip()
            )

            # โหลด Reference File
            df_ref = pd.read_excel("data/Client.xlsx")
            df_ref.columns = (
                df_ref.columns.astype(str)
                .str.encode('ascii', 'ignore').str.decode('utf-8')
                .str.replace(r'\s+', ' ', regex=True)
                .str.strip()
            )

            # ตรวจสอบคอลัมน์ใน reference
            required_ref_cols = {
                "Mapping", "Maximum threshold(out)", "Minimum threshold(out)",
                "Maximum threshold(in)", "Minimum threshold(in)"
            }
            if not required_ref_cols.issubset(df_ref.columns):
                st.error(f"Reference file must contain columns: {', '.join(required_ref_cols)}")
                st.stop()

            df_ref["Mapping"] = df_ref["Mapping"].astype(str).str.strip()

            col_out = "Output Optical Power (dBm)"
            col_in = "Input Optical Power(dBm)"
            col_max_out = "Maximum threshold(out)"
            col_min_out = "Minimum threshold(out)"
            col_max_in = "Maximum threshold(in)"
            col_min_in = "Minimum threshold(in)"

            # ✅ เพิ่ม order ใน reference เพื่อคุมลำดับแสดงผลตามไฟล์ ref
            df_ref["order"] = range(len(df_ref))

            # Merge ข้อมูล (พก order มาด้วย)
            df_merged = pd.merge(
                df_client,
                df_ref[["Site Name", "Mapping", col_max_out, col_min_out, col_max_in, col_min_in, "order"]],
                left_on="Mapping Format",
                right_on="Mapping",
                how="inner"
            )

            # ตรวจสอบผลลัพธ์
            if not df_merged.empty:
                # ✅ เรียงตาม order จาก ref แล้วจัดคอลัมน์ที่จะแสดง
                df_merged = df_merged.sort_values("order").reset_index(drop=True)

                df_result = df_merged[[
                    "Site Name", "ME", "Measure Object",
                    col_max_out, col_min_out, col_out,
                    col_max_in,  col_min_in, col_in
                ]].copy()

                # ✅ แปลงตัวเลขให้ชัวร์ก่อนใช้เทียบ
                for c in [col_out, col_in, col_max_out, col_min_out, col_max_in, col_min_in]:
                    if c in df_result.columns:
                        df_result.loc[:, c] = pd.to_numeric(df_result[c], errors="coerce")

                # --------- เพิ่ม Filter แบบไล่ชั้น (ให้เหมือนเมนูอื่น) ---------
                df_filtered, _sel = cascading_filter(
                    df_result,
                    cols=["Site Name", "ME", "Measure Object"],
                    ns="client",
                    clear_text="Clear Client Filters"
                )
                st.caption(f"Client (showing {len(df_filtered)}/{len(df_result)} rows)")

                # --------- ไฮไลท์เหมือนเมนูอื่น: เทาทั้งแถว + แดงเฉพาะค่าที่ผิด ---------
                def row_has_issue(r):
                    try:
                        return (
                            (r[col_out] > r[col_max_out]) or (r[col_out] < r[col_min_out]) or
                            (r[col_in]  > r[col_max_in])  or (r[col_in]  < r[col_min_in])
                        )
                    except Exception:
                        return False

                def highlight_critical_cells(val, colname, row):
                    try:
                        if colname == col_out:
                            return 'background-color:#ff4d4d; color:white' if (val > row[col_max_out] or val < row[col_min_out]) else ''
                        elif colname == col_in:
                            return 'background-color:#ff4d4d; color:white' if (val > row[col_max_in] or val < row[col_min_in]) else ''
                        return ''
                    except Exception:
                        return ''

                df_view = df_filtered.copy()

                styled_df = (
                    df_view.style
                    # เทาทั้งแถวเมื่อมีปัญหา
                    .apply(lambda r: ['background-color:#e6e6e6; color:black' if row_has_issue(r) else '' for _ in r], axis=1)
                    # แดงเฉพาะค่าที่ผิด (ทั้ง out/in)
                    .apply(lambda row: [highlight_critical_cells(row[colname], colname, row) for colname in df_view.columns], axis=1)
                    .format({
                        col_max_out: "{:.2f}",
                        col_min_out: "{:.2f}",
                        col_max_in:  "{:.2f}",
                        col_min_in:  "{:.2f}",
                        col_out:     "{:.2f}",
                        col_in:      "{:.2f}",
                    })
                )

                st.markdown("### Client Performance")
                st.write(styled_df, use_container_width=True)

                # --------- แบนเนอร์สถานะ (สอดคล้องกับเมนูอื่น) ---------
                failed_rows = df_view.apply(row_has_issue, axis=1)
                st.markdown(
                    "<div style='text-align:center; font-size:32px; font-weight:bold; color:{};'>Client Performance {}</div>".format(
                        "red" if failed_rows.any() else "green",
                        "Warning" if failed_rows.any() else "Normal"
                    ),
                    unsafe_allow_html=True
                )

            else:
                st.warning("No matching mapping found between Client file and reference")

        except Exception as e:
            st.error(f"An error occurred during processing: {e}")
    else:
        st.info("Please upload a Client file to start the analysis")


elif menu == "Line board":
    st.markdown("### Upload Line cards performance File")

    # ใช้ key แบบสั้นเฉพาะเมนูนี้
    uploaded_line = st.file_uploader("Upload Line cards File", type=["xlsx"], key="lb_line")
    uploaded_log  = st.file_uploader("Upload WASON Log", type=["txt"], key="lb_log")

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

    # แคช pmap จาก log (ถ้ามี)
    if uploaded_log:
        log_text = uploaded_log.read().decode("utf-8", errors="ignore")
        st.session_state.lb_pmap = get_preset_map(log_text)

    pmap = st.session_state.get("lb_pmap", {})  # ใช้ที่แคชไว้ถ้าไม่มี log รอบนี้

    # แคชไฟล์ line (เก็บทั้ง DataFrame และชื่อไฟล์)
    if uploaded_line:
        st.session_state.lb_data = pd.read_excel(uploaded_line)
        st.session_state.lb_file = uploaded_line.name
        st.success(f"Line cards file loaded: {st.session_state.lb_file}")
    elif st.session_state.get("lb_file"):
        st.info(f"Using cached data: {st.session_state.lb_file}")

    # มีข้อมูลแล้วค่อยประมวลผล (ไม่บังคับมี log)
    if st.session_state.get("lb_data") is not None:

        try:
            df_line = st.session_state.lb_data.copy()
            df_line.columns = (
                df_line.columns.str.strip()
                .str.replace(r'\s+', ' ', regex=True)
                .str.replace('\u00a0', ' ')
            )

            df_ref = pd.read_excel("data/Line.xlsx")
            df_ref.columns = (
                df_ref.columns.astype(str)
                .str.replace(r'\s+', ' ', regex=True)
                .str.replace('\u00a0', ' ')
                .str.strip()
            )

            # เพิ่มลำดับไว้ใช้เรียงภายหลัง
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

                # เรียงตามลำดับจาก ref แล้วลบคอลัมน์ order ทิ้ง
                df_result = df_result.sort_values("order").drop(columns=["order"]).reset_index(drop=True)

                # ---------- (1) FILTER: ใช้ cascading_filter ----------
                df_filtered, _sel = cascading_filter(
                    df_result,
                    cols=["Site Name", "ME", "Measure Object","Call ID","Route"],
                    ns="line",
                    clear_text="Clear Line Filters"
                )
                st.caption(f"Line Performance (showing {len(df_filtered)}/{len(df_result)} rows)")

                # ---------- เตรียมตัวเลข ----------
                num_cols = ["Instant BER After FEC", col_out, col_in, col_max_out, col_min_out, col_max_in, col_min_in]
                num_cols = [c for c in num_cols if c in df_filtered.columns]

                df_view = df_filtered.copy()
                df_view.loc[:, num_cols] = df_view[num_cols].apply(pd.to_numeric, errors="coerce")

                # ---------- ตรรกะตรวจปัญหา ----------
                def row_has_issue(r):
                    return (
                        (pd.notna(r.get("Instant BER After FEC")) and float(r["Instant BER After FEC"]) > 0) or
                        (pd.notna(r.get(col_out)) and pd.notna(r.get(col_max_out)) and r[col_out] > r[col_max_out]) or
                        (pd.notna(r.get(col_out)) and pd.notna(r.get(col_min_out)) and r[col_out] < r[col_min_out]) or
                        (pd.notna(r.get(col_in)) and pd.notna(r.get(col_max_in)) and r[col_in] > r[col_max_in]) or
                        (pd.notna(r.get(col_in)) and pd.notna(r.get(col_min_in)) and r[col_in] < r[col_min_in])
                    )

                # ---------- (3) HIGHLIGHT: เทาทั้งแถว + แดงเฉพาะช่อง + ฟ้า Route ----------
                styled = (
                    df_view.style
                    # เทาทั้งแถวถ้ามีปัญหา
                    .apply(lambda r: ['background-color:#e6e6e6; color:black' if row_has_issue(r) else '' for _ in r], axis=1)
                    # แดง BER
                    .apply(lambda _: [
                        'background-color:#ff4d4d; color:white' if (pd.notna(v) and float(v) > 0) else ''
                        for v in df_view["Instant BER After FEC"]
                    ], subset=["Instant BER After FEC"])
                    # แดง Output Power
                    .apply(lambda _: [
                        'background-color:#ff4d4d; color:white' if (pd.notna(v) and pd.notna(hi) and pd.notna(lo) and (v > hi or v < lo)) else ''
                        for v, hi, lo in zip(df_view[col_out], df_view[col_max_out], df_view[col_min_out])
                    ], subset=[col_out])
                    # แดง Input Power
                    .apply(lambda _: [
                        'background-color:#ff4d4d; color:white' if (pd.notna(v) and pd.notna(hi) and pd.notna(lo) and (v > hi or v < lo)) else ''
                        for v, hi, lo in zip(df_view[col_in], df_view[col_max_in], df_view[col_min_in])
                    ], subset=[col_in])
                    # ฟ้า Route ที่เป็น Preset
                    .apply(lambda _: [
                        'background-color:lightblue; color:black' if str(x).startswith("Preset") else ''
                        for x in df_view["Route"]
                    ], subset=["Route"])
                    .format({
                        col_out: "{:.2f}", col_in: "{:.2f}",
                        col_max_out: "{:.0f}", col_min_out: "{:.0f}",
                        col_max_in: "{:.0f}", col_min_in: "{:.0f}",
                        "Instant BER After FEC": "{:.1e}"
                    })
                )
                st.markdown("### Line Performance")
                st.write(styled)

                # สรุปสถานะอิงข้อมูลที่ถูกฟิลเตอร์แล้ว (คงเดิม)
                failed_rows = df_view.apply(row_has_issue, axis=1)
                st.markdown(
                    "<div style='text-align:center; font-size:32px; font-weight:bold; color:{};'>Line Performance {}</div>".format(
                        "red" if failed_rows.any() else "green",
                        "Warning" if failed_rows.any() else "Normal"
                    ),
                    unsafe_allow_html=True
                )
            else:
                st.warning("No matching mapping found between Line file and reference")
        except Exception as e: 
            st.error(f"An error occurred during processing: {e}")
    else:
        st.info("Please upload a Line file first")


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

    # Process when both files are uploaded
    if "osc_optical_data" in st.session_state and "osc_fm_data" in st.session_state:
        try:
            df_optical = st.session_state.osc_optical_data.copy()
            df_fm = st.session_state.osc_fm_data.copy()

            df_optical.columns = df_optical.columns.str.strip()
            df_optical["Max - Min (dB)"] = (
                df_optical["Max Value of Input Optical Power(dBm)"]
                - df_optical["Min Value of Input Optical Power(dBm)"]
            )

            # Extract Target ME (kept for logic; will not display)
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

            # Find no-match
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
                # Cascading filter (show only highlighted table)
                df_nomatch_filtered, _sel = cascading_filter(
                    df_nomatch,
                    cols=["ME", "Measure Object"],
                    ns="fiber",
                    labels={"ME": "Managed Element"},
                    clear_text="Clear Fiber Filters"
                )
                st.caption(f"Fiber Flapping (showing {len(df_nomatch_filtered)}/{len(df_nomatch)} rows)")

                # Columns to display (no Target ME)
                view_cols = [
                    "Begin Time", "End Time", "Granularity", "ME", "ME IP", "Measure Object",
                    "Max Value of Input Optical Power(dBm)", "Min Value of Input Optical Power(dBm)",
                    "Input Optical Power(dBm)", "Max - Min (dB)"
                ]
                view_cols = [c for c in view_cols if c in df_nomatch_filtered.columns]
                df_view = df_nomatch_filtered[view_cols].copy()

                # Cast numerics for formatting (2 decimals)
                num_cols = [
                    "Max Value of Input Optical Power(dBm)",
                    "Min Value of Input Optical Power(dBm)",
                    "Input Optical Power(dBm)",
                    "Max - Min (dB)"
                ]
                num_cols = [c for c in num_cols if c in df_view.columns]
                if num_cols:
                    df_view.loc[:, num_cols] = df_view[num_cols].apply(pd.to_numeric, errors="coerce")

                # Highlight ONLY red on "Max - Min (dB)" > 2 (no gray rows)
                styled_df = (
                    df_view.style
                    .apply(
                        lambda _:
                            ['background-color:#ff4d4d; color:white' if (v > 2) else '' 
                             for v in df_view["Max - Min (dB)"]],
                        subset=["Max - Min (dB)"]
                    )
                    .format({c: "{:.2f}" for c in num_cols})
                )

                # Show only the highlighted table
                st.write(styled_df, use_container_width=True)

                # Daily graph (based on filtered view)
                df_view["Date"] = pd.to_datetime(df_view["Begin Time"]).dt.date
                site_count = df_view.groupby("Date")["ME"].nunique().reset_index()
                site_count.columns = ["Date", "Sites"]

                import plotly.express as px
                fig = px.bar(site_count, x="Date", y="Sites", text="Sites", title="No Fiber Break Alarm Match")
                fig.update_traces(textposition="outside")
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

                # Daily detail (still no Target ME)
                for date, group in df_view.groupby("Date"):
                    st.markdown(f"#### {date.strftime('%Y-%m-%d')}")
                    st.dataframe(group[view_cols], use_container_width=True)
            else:
                st.success("No unmatched fiber flapping records found")

        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.info("Please upload both OSC and FM files first")


elif menu == "Preset status":
    st.subheader("Preset status")
    st.caption("Upload a MobaXterm log and review only CALLs that are on WR.")

    # simple CSS (ปลอดภัยจะใส่ใน elif ได้)
    st.markdown("""
    <style>
      .kpi-row { display:grid; grid-template-columns: repeat(3,1fr); gap:.75rem; margin:.25rem 0 1rem 0;}
      .kpi-card { border:1px solid rgba(0,0,0,.06); background:#fafafa; border-radius:14px; padding:12px 14px;}
      .kpi-card .label{font-size:12px;color:#6b7280;} .kpi-card .value{font-size:24px;font-weight:700;margin-top:2px;}
      .pill{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;margin-right:6px;border:1px solid transparent;}
      .pill-green{background:#ecfdf5;border-color:#10b98144;} .pill-blue{background:#eff6ff;border-color:#3b82f633;}
      .pill-red{background:#fef2f2;border-color:#ef444444;}
    </style>
    """, unsafe_allow_html=True)

    up = st.file_uploader("Upload MobaXterm log (.txt)", type=["txt"], key="preset_file")
    if not up:
        st.info("Drop a log file to analyze.")
        st.stop()

    text = up.read().decode("utf-8", errors="ignore")
    calls = parse_calls(text)  # <- helper defined above

    # keep only calls with WR
    rows = []
    for cb in calls:
        res = evaluate_preset_status(cb)  # <- helper defined above
        if res.get("has_wr"):
            rows.append({
                "Call": cb.call_id,
                "Triple": cb.triple,
                "Preroute": res.get("pr_index"),
                "Verdict": res.get("verdict"),
                "Restore": res.get("Restore"),
                "Raw": res.get("raw"),
            })

    if not rows:
        st.warning("No CALLs with WR were found in this file.")
        st.stop()

    df = pd.DataFrame(rows).sort_values("Call").reset_index(drop=True)
    total = len(df); passes = int((df["Verdict"] == "PASS").sum()); fails = int((df["Verdict"] == "FAIL").sum())

    # KPIs
    c1, c2, c3 = st.columns(3)
    with c1: st.markdown(f'<div class="kpi-card"><div class="label">Total WR calls</div><div class="value">{total}</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><div class="label">Pass</div><div class="value">{passes}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><div class="label">Abnormal</div><div class="value">{fails}</div></div>', unsafe_allow_html=True)

    # controls
    left, right = st.columns([1,1])
    with left:
        only_abnormal = st.checkbox("Show only abnormal", value=False)
    with right:
        st.download_button(
            "Download summary (CSV)",
            df.drop(columns=["Raw"]).to_csv(index=False).encode("utf-8"),
            file_name="preset_summary.csv",
            mime="text/csv",
            use_container_width=True
        )

    view = df if not only_abnormal else df[df["Verdict"] == "FAIL"]
    st.dataframe(view.drop(columns=["Raw"]), use_container_width=True, hide_index=True)

    # per-call cards
    for _, r in view.iterrows():
        with st.container(border=True):
            st.markdown(f"**Call {int(r.Call)}** · `{r.Triple}`")
            if r.Verdict == "PASS":
                st.success("normal")
                st.write(f"Call {int(r.Call)} [{r.Triple}] uses preroute #{int(r.Preroute)}")
                st.markdown(
                    '<span class="pill pill-green">WR NO_ALARM</span>'
                    f'<span class="pill pill-blue">Preroute #{int(r.Preroute)}</span>'
                    '<span class="pill pill-green">SUCCESS</span>',
                    unsafe_allow_html=True
                )
            else:
                st.error("Abnormal")
                st.write(str(r.Restore)) 
                if pd.notna(r.Preroute):
                    st.markdown(f'<span class="pill pill-blue">Preroute #{int(r.Preroute)}</span>', unsafe_allow_html=True)

            with st.expander("Show raw log"):
                st.code(str(r.Raw), language="text")

                

# region Loss between EOL
elif menu == "Loss between EOL":
    st.markdown("### Please upload files")

    uploaded_raw_eol = st.file_uploader("Upload Raw Optical Attenuation", type=["xlsx"], key="raw_optical_atten")
    if uploaded_raw_eol:
        df_raw_data = pd.read_excel(uploaded_raw_eol)

        st.session_state.raw_eol_data = df_raw_data
        st.success("Raw Data File Uploaded")

    analyzer = EOLAnalyzer(
        st.session_state.get("reference_sheet"), 
        st.session_state.get("raw_eol_data"),
    )

    analyzer.process()
    
# region Loss Between Core
elif menu == "Loss between Core":
    st.markdown("### Please upload files")

    uploaded_raw_eol = st.file_uploader("Upload Raw Optical Attenuation", type=["xlsx"], key="raw_optical_atten")
    if uploaded_raw_eol:
        df_raw_data = pd.read_excel(uploaded_raw_eol)

        st.session_state.raw_eol_data = df_raw_data
        st.success("Raw Data File Uploaded")

    analyzer = CoreAnalyzer(
        st.session_state.get("reference_sheet"), 
        st.session_state.get("raw_eol_data"),
    )

    analyzer.process()

elif menu == "Reference Sheet":
    st.markdown("### Reference Sheet")

    st.dataframe(st.session_state.get("reference_sheet"), height=700)