
import io
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

# ============================================================
# LABSTREAM AUTOMATION — SMART SOLUTION PREP
# Software prototype / simulation
# ============================================================

st.set_page_config(
    page_title="Smart Solution Prep | LabStream",
    page_icon="🧪",
    layout="wide"
)

# ============================================================
# LABSTREAM AUTOMATION — BRAND DESIGN
# ============================================================

st.markdown("""
<style>
    /* Main background */
    [data-testid="stAppViewContainer"] {
        background-color: #0B1930;
        color: #F4F8FC;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #102842;
    }

    /* Header */
    [data-testid="stHeader"] {
        background-color: #0B1930;
    }

    /* Main headings */
    h1, h2, h3 {
        color: #F4F8FC;
    }

    /* Subtle accent for dividers */
    hr {
        border-color: #31516B;
    }
</style>
""", unsafe_allow_html=True)

INSTRUMENT_ID = "FHNW-LCMS-001"
STATION_ID = "SSP-001"
RECIPE_VERSION = "MPA-V1.0"

# Ammonium formate molar mass (g/mol)
MOLAR_MASS_AF = 63.06

LOGO_PATH = Path(__file__).parent / "labstream_logo.png"

# ============================================================
# SESSION STATE
# ============================================================

DEFAULTS = {
    "stock_ml": 350.0,
    "completed_runs": 0,
    "runs_since_cleaning": 0,
    "cleaning_interval": 5,
    "batches": [],
    "run_logs": [],
    "service_logs": [],
    "batch_counter": 1,
    "current_batch": None,
}

for name, value in DEFAULTS.items():
    if name not in st.session_state:
        st.session_state[name] = (
            value.copy() if isinstance(value, list) else value
        )


def now():
    return datetime.now().isoformat(timespec="seconds")


# ============================================================
# PAGE HEADER WITH LOGO
# ============================================================

def page_header(title, subtitle):
    left, right = st.columns([5, 1.2])

    with left:
        st.title(title)
        st.caption(subtitle)

    with right:
        if LOGO_PATH.is_file():
            st.image(str(LOGO_PATH), use_container_width=True)
        else:
            st.markdown("### LabStream Automation")


# ============================================================
# RECORD FUNCTIONS
# ============================================================

def store_batch(batch):
    records = st.session_state.batches

    for i, record in enumerate(records):
        if record["batch_id"] == batch["batch_id"]:
            records[i] = batch.copy()
            return

    records.append(batch.copy())


def flatten_batch(batch):
    return {
        "Batch ID": batch["batch_id"],
        "Created": batch["created"],
        "Instrument": batch["instrument"],
        "Recipe": batch["recipe"],
        "Recipe version": batch["recipe_version"],
        "Target volume (mL)": batch["target_volume_ml"],
        "Ammonium formate target (g)": batch["target_mass_g"],
        "Formic acid target (mL)": batch["target_acid_ml"],
        "Ammonium formate actual (g)": batch.get("actual_mass_g"),
        "Formic acid actual (mL)": batch.get("actual_acid_ml"),
        "Final volume actual (mL)": batch.get("actual_volume_ml"),
        "pH": batch.get("ph"),
        "Conductivity (mS/cm)": batch.get(
            "conductivity_ms_cm"
        ),
        "Status": batch["status"],
        "Transferred": batch.get("transferred_at", ""),
    }


# ============================================================
# EXCEL EXPORT
# ============================================================

def workbook_bytes():
    wb = Workbook()
    wb.remove(wb.active)

    sheets = {
        "Batches": [
            flatten_batch(b)
            for b in st.session_state.batches
        ],

        "Runs": st.session_state.run_logs,

        "Service": [
            {
                "Timestamp": r["timestamp"],
                "Operator": r["operator"],
                "Cleaning duration (min)":
                    r["cleaning_duration_min"],
                "Cleaning status": r["cleaning_status"],
                "Maintenance": r["maintenance_status"],
                "Calibration": r["calibration_status"],
                "Runs before cleaning":
                    r["runs_before_cleaning"],
                "Notes": r["notes"],
            }
            for r in st.session_state.service_logs
        ],

        "Material lots": [
            {
                "Batch ID": b["batch_id"],
                "Ammonium formate lot":
                    b["material_lots"]["ammonium_formate"],
                "Formic acid lot":
                    b["material_lots"]["formic_acid"],
                "Water lot":
                    b["material_lots"]["water"],
            }
            for b in st.session_state.batches
        ],

        "QC": [
            {
                "Batch ID": b["batch_id"],
                **check
            }
            for b in st.session_state.batches
            for check in b.get("qc_checks", [])
        ],
    }

    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(sheet_name)

        if not rows:
            ws.append(["No records in current session"])
            continue

        columns = list(rows[0].keys())
        ws.append(columns)

        for row in rows:
            ws.append([
                str(row.get(c, ""))
                if isinstance(row.get(c), (dict, list))
                else row.get(c)
                for c in columns
            ])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )
            cell.fill = PatternFill(
                "solid",
                fgColor="184C68"
            )

        for column in ws.columns:
            letter = column[0].column_letter
            width = min(
                52,
                max(
                    15,
                    max(
                        len(str(cell.value or ""))
                        for cell in column
                    ) + 2
                )
            )
            ws.column_dimensions[letter].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# PDF EXPORT
# ============================================================

def pdf_bytes():
    output = io.BytesIO()

    doc = SimpleDocTemplate(
        output,
        pagesize=(595, 842),
        leftMargin=40,
        rightMargin=40
    )

    styles = getSampleStyleSheet()

    story = [
        Paragraph(
            "LabStream Automation | Smart Solution Prep",
            styles["Title"]
        ),
        Paragraph(
            "SIMULATED laboratory documentation",
            styles["Normal"]
        ),
        Spacer(1, 16),
        Paragraph(
            f"Instrument: {INSTRUMENT_ID}",
            styles["Normal"]
        ),
        Paragraph(
            f"Preparation station: {STATION_ID}",
            styles["Normal"]
        ),
        Paragraph(
            f"Export: {now()}",
            styles["Normal"]
        ),
        Spacer(1, 12),
    ]

    sections = [
        (
            "Batch history",
            [
                flatten_batch(b)
                for b in st.session_state.batches
            ]
        ),
        (
            "Simulated run history",
            st.session_state.run_logs
        ),
        (
            "Service history",
            [
                {
                    "Time": r["timestamp"],
                    "Operator": r["operator"],
                    "Cleaning": r["cleaning_status"],
                    "Maintenance": r["maintenance_status"],
                    "Calibration": r["calibration_status"],
                    "Duration (min)":
                        r["cleaning_duration_min"],
                }
                for r in st.session_state.service_logs
            ]
        ),
    ]

    for heading, records in sections:
        story.append(
            Paragraph(heading, styles["Heading2"])
        )

        if not records:
            story.append(
                Paragraph("No records.", styles["Normal"])
            )

        else:
            for record in records:
                for key, value in record.items():
                    safe_key = str(key).replace(
                        "&", "&amp;"
                    ).replace("<", "&lt;")

                    safe_value = str(value).replace(
                        "&", "&amp;"
                    ).replace("<", "&lt;")

                    story.append(
                        Paragraph(
                            f"<b>{safe_key}:</b> {safe_value}",
                            styles["BodyText"]
                        )
                    )

                story.append(Spacer(1, 8))

        story.append(Spacer(1, 10))

    doc.build(story)
    return output.getvalue()


# ============================================================
# SIDEBAR — NAVIGATION
# ============================================================

st.sidebar.markdown("## 🧪 LabStream")
st.sidebar.caption(
    "Smart Solution Prep | Software Prototype"
)

page = st.sidebar.radio(
    "Navigation",
    [
        "🏠 Dashboard",
        "🧼 Reinigung & Wartung",
        "📋 Dokumentation"
    ]
)

st.sidebar.divider()

st.sidebar.selectbox(
    "Cleaning interval (simulated runs)",
    [5, 10, 15],
    key="cleaning_interval",
    help=(
        "Demonstration threshold only. "
        "Not a validated laboratory cleaning interval."
    )
)

st.sidebar.caption(
    f"Virtual instrument: {INSTRUMENT_ID}"
)

st.sidebar.warning(
    "SIMULATION — No physical hardware connection"
)
# ============================================================
# PROJECT TEAM — LABSTREAM AUTOMATION
# ============================================================

st.sidebar.divider()

st.sidebar.markdown(
    "<span style='color:#7AD8DB; font-weight:700;'>"
    "PROJECT TEAM</span>",
    unsafe_allow_html=True
)

st.sidebar.caption("FHNW Life Sciences | 2026")

st.sidebar.markdown("""
<div style="line-height:1.9; color:#F4F8FC; font-size:14px;">
    Yasmine Ben Sayed<br>
    Irina Brandenberger<br>
    Esma Kolbasi<br>
    Hajar Maânaoui
</div>
""", unsafe_allow_html=True)

st.sidebar.caption(
    "LabStream Automation | Smart Solution Prep"
)
cleaning_due = (
    st.session_state.runs_since_cleaning
    >= st.session_state.cleaning_interval
)


# ============================================================
# PAGE 1 — DASHBOARD
# ============================================================

if page == "🏠 Dashboard":

    page_header(
        "🧪 Smart Solution Prep",
        "Automated LC-MS Mobile Phase Management | Simulation"
    )

    # --------------------------------------------------------
    # 1. INSTRUMENT
    # --------------------------------------------------------

    st.header("1. Instrument & Mobile Phase")

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "LC-MS Instrument",
        INSTRUMENT_ID
    )

    c2.metric(
        "Connection",
        "VIRTUAL"
    )

    c3.metric(
        "Reservoir",
        "RES-A-001"
    )

    st.info(
        "Mobile Phase A — Default recipe: "
        "5 mM ammonium formate + "
        "0.1% formic acid (v/v)"
    )

    st.divider()

    # --------------------------------------------------------
    # 2. RUN PLANNING
    # --------------------------------------------------------

    st.header("2. Run Planning & Inventory")

    c1, c2, c3 = st.columns(3)

    with c1:
        flow = st.number_input(
            "Flow rate (mL/min)",
            min_value=0.01,
            value=0.5,
            step=0.1
        )

    with c2:
        duration = st.number_input(
            "Run duration (min)",
            min_value=1.0,
            value=20.0
        )

    with c3:
        runs = st.number_input(
            "Planned runs",
            min_value=1,
            value=40
        )

    reserve = st.number_input(
        "Safety reserve (mL)",
        min_value=0.0,
        value=50.0
    )

    consumption = flow * duration * runs

    required = consumption + reserve

    shortage = max(
        0.0,
        required - st.session_state.stock_ml
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Available mobile phase",
        f"{st.session_state.stock_ml:.1f} mL"
    )

    c2.metric(
        "Planned consumption",
        f"{consumption:.1f} mL"
    )

    c3.metric(
        "Required incl. reserve",
        f"{required:.1f} mL"
    )

    st.progress(
        min(
            st.session_state.stock_ml
            / max(required, 0.001),
            1.0
        )
    )

    if shortage > 0:
        st.error(
            "LOW MOBILE PHASE — Preparation needed. "
            f"Shortage: {shortage:.1f} mL"
        )
    else:
        st.success(
            "Inventory sufficient for planned runs "
            "and safety reserve."
        )

    st.caption(
        "Simplified single-phase consumption model. "
        "Gradient composition, priming and flushing "
        "are not included."
    )

    st.divider()

    # --------------------------------------------------------
    # 3. AUTOMATED PREPARATION
    # --------------------------------------------------------

    st.header(
        "3. Automated Preparation — Dynamic Recipe"
    )

    c1, c2 = st.columns(2)

    with c1:

        concentration_mm = st.number_input(
            "Ammonium formate target (mM)",
            min_value=0.1,
            value=5.0,
            step=0.5
        )

        acid_percent = st.number_input(
            "Formic acid target (% v/v)",
            min_value=0.0,
            max_value=100.0,
            value=0.1,
            step=0.1,
            format="%.2f"
        )

    with c2:

        target_volume = st.number_input(
            "Final preparation volume (mL)",
            min_value=1.0,
            value=1000.0,
            step=100.0
        )

        acid_stock_percent = st.number_input(
            "Formic acid stock (% v/v)",
            min_value=0.01,
            max_value=100.0,
            value=100.0
        )

    # Dynamic calculation
    target_mass = (
        concentration_mm
        * MOLAR_MASS_AF
        * target_volume
        / 1_000_000
    )

    target_acid = (
        acid_percent
        / acid_stock_percent
        * target_volume
    )

    valid_recipe = (
        target_acid <= target_volume
    )

    recipe = (
        f"{concentration_mm:g} mM ammonium formate "
        f"+ {acid_percent:g}% formic acid (v/v)"
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Ammonium formate",
        f"{target_mass:.4f} g"
    )

    c2.metric(
        "Formic acid stock",
        f"{target_acid:.3f} mL"
    )

    c3.metric(
        "Final solution volume",
        f"{target_volume:.1f} mL"
    )

    st.caption(
        "LC-MS water: add to the final volume."
    )

    st.caption(
        "The stock concentration is treated as % v/v. "
        "For commercial acids specified in % w/w, "
        "density and purity must be considered."
    )

    if not valid_recipe:
        st.error(
            "Invalid recipe: calculated acid stock volume "
            "exceeds final preparation volume."
        )

    # Material lots
    with st.expander(
        "Material lot IDs — Simulation"
    ):

        lot_af = st.text_input(
            "Ammonium formate lot",
            "AF-LOT-2401"
        )

        lot_fa = st.text_input(
            "Formic acid lot",
            "FA-LOT-2402"
        )

        lot_water = st.text_input(
            "LC-MS water lot",
            "W-LOT-2403"
        )

    active_batch = st.session_state.current_batch

    unfinished = (
        active_batch is not None
        and not active_batch.get("transferred_at")
    )

    if st.button(
        "▶ Simulate preparation",
        type="primary",
        disabled=(
            not valid_recipe
            or shortage <= 0
            or unfinished
        )
    ):

        batch_id = (
            f"MPA-{datetime.now():%Y%m%d}-"
            f"{st.session_state.batch_counter:03d}"
        )

        st.session_state.batch_counter += 1

        batch = {
            "batch_id": batch_id,
            "created": now(),
            "instrument": INSTRUMENT_ID,
            "station": STATION_ID,
            "recipe": recipe,
            "recipe_version": RECIPE_VERSION,
            "target_concentration_mm":
                concentration_mm,
            "target_acid_percent_vv":
                acid_percent,
            "acid_stock_percent_vv":
                acid_stock_percent,
            "target_volume_ml": target_volume,
            "target_mass_g": target_mass,
            "target_acid_ml": target_acid,
            "material_lots": {
                "ammonium_formate": lot_af,
                "formic_acid": lot_fa,
                "water": lot_water
            },
            "status": "PENDING QC",
            "transferred_at": None,
            "actual_mass_g": None,
            "actual_acid_ml": None,
            "actual_volume_ml": None,
            "ph": None,
            "conductivity_ms_cm": None,
            "qc_checks": []
        }

        st.session_state.current_batch = batch

        store_batch(batch)

        st.rerun()

    # --------------------------------------------------------
    # 4. QC
    # --------------------------------------------------------

    batch = st.session_state.current_batch

    if batch is not None:

        st.info(
            f"Current batch: {batch['batch_id']} "
            f"| {batch['recipe']} "
            f"| Status: {batch['status']}"
        )

        st.header(
            "4. Dosing & Quality Control"
        )

        st.caption(
            "Simulated or manually entered measurements. "
            "Tolerances are illustrative engineering values, "
            "not validated LC-MS specifications."
        )

        c1, c2, c3 = st.columns(3)

        bid = batch["batch_id"]

        with c1:
            actual_mass = st.number_input(
                "Ammonium formate actual (g)",
                min_value=0.0,
                value=round(
                    batch["target_mass_g"], 4
                ),
                format="%.4f",
                key=f"mass_{bid}"
            )

        with c2:
            actual_acid = st.number_input(
                "Formic acid stock actual (mL)",
                min_value=0.0,
                value=round(
                    batch["target_acid_ml"], 3
                ),
                format="%.3f",
                key=f"acid_{bid}"
            )

        with c3:
            actual_volume = st.number_input(
                "Final volume actual (mL)",
                min_value=0.0,
                value=batch["target_volume_ml"],
                format="%.2f",
                key=f"volume_{bid}"
            )

        st.subheader("pH & Conductivity")

        c1, c2 = st.columns(2)

        with c1:

            ph_actual = st.number_input(
                "Measured pH (simulated)",
                min_value=0.0,
                max_value=14.0,
                value=3.2,
                key=f"ph_{bid}"
            )

            ph_target = st.number_input(
                "Reference pH (illustrative)",
                min_value=0.0,
                max_value=14.0,
                value=3.2
            )

            ph_tolerance = st.number_input(
                "Allowed pH difference",
                min_value=0.0,
                value=0.2
            )

        with c2:

            conductivity_actual = st.number_input(
                "Measured conductivity (mS/cm)",
                min_value=0.0,
                value=0.45,
                key=f"conductivity_{bid}"
            )

            conductivity_target = st.number_input(
                "Reference conductivity (mS/cm)",
                min_value=0.001,
                value=0.45
            )

            conductivity_tolerance = st.number_input(
                "Conductivity tolerance (%)",
                min_value=0.0,
                value=10.0
            )

        def percent_deviation(actual, target):
            if target == 0:
                return (
                    0.0 if actual == 0
                    else float("inf")
                )

            return (
                abs(actual - target)
                / target * 100
            )

        check_data = [
            (
                "Ammonium formate",
                percent_deviation(
                    actual_mass,
                    batch["target_mass_g"]
                ),
                1.0,
                "%"
            ),
            (
                "Formic acid stock",
                percent_deviation(
                    actual_acid,
                    batch["target_acid_ml"]
                ),
                1.0,
                "%"
            ),
            (
                "Final volume",
                percent_deviation(
                    actual_volume,
                    batch["target_volume_ml"]
                ),
                0.5,
                "%"
            ),
            (
                "pH",
                abs(ph_actual - ph_target),
                ph_tolerance,
                "pH units"
            ),
            (
                "Conductivity",
                percent_deviation(
                    conductivity_actual,
                    conductivity_target
                ),
                conductivity_tolerance,
                "%"
            )
        ]

        qc_checks = [
            {
                "Parameter": name,
                "Deviation": round(deviation, 4),
                "Limit": limit,
                "Unit": unit,
                "Result": (
                    "PASS"
                    if deviation <= limit
                    else "FAIL"
                )
            }
            for name, deviation, limit, unit
            in check_data
        ]

        st.dataframe(
            pd.DataFrame(qc_checks),
            hide_index=True,
            use_container_width=True
        )

        if st.button(
            "Evaluate QC & save batch",
            disabled=(
                batch.get("transferred_at")
                is not None
            )
        ):

            batch.update({
                "actual_mass_g": actual_mass,
                "actual_acid_ml": actual_acid,
                "actual_volume_ml": actual_volume,
                "ph": ph_actual,
                "conductivity_ms_cm":
                    conductivity_actual,
                "qc_checks": qc_checks,
                "status": (
                    "SIMULATED RELEASE"
                    if all(
                        c["Result"] == "PASS"
                        for c in qc_checks
                    )
                    else "BLOCKED"
                ),
                "qc_checked_at": now()
            })

            store_batch(batch)
            st.rerun()

        if batch["status"] == "BLOCKED":

            st.error(
                "QC FAILED — BATCH BLOCKED"
            )

        elif batch["status"] == "SIMULATED RELEASE":

            st.success(
                "SIMULATED QC RELEASE"
            )

            if st.button(
                "Simulate transfer to RES-B-001"
            ):

                st.session_state.stock_ml += (
                    batch["target_volume_ml"]
                )

                batch["transferred_at"] = now()

                batch["status"] = (
                    "TRANSFERRED (SIMULATED)"
                )

                store_batch(batch)
                st.rerun()

        elif batch["status"] == (
            "TRANSFERRED (SIMULATED)"
        ):

            st.success(
                "Virtual transfer completed. "
                "Inventory updated."
            )

        else:

            st.warning(
                "Batch awaiting QC."
            )

    st.divider()

    # --------------------------------------------------------
    # 5. RUN SIMULATION
    # --------------------------------------------------------

    st.header(
        "5. Run Simulation & Cleaning Alert"
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Completed simulated runs",
        st.session_state.completed_runs
    )

    c2.metric(
        "Runs since cleaning",
        (
            f"{st.session_state.runs_since_cleaning}"
            f"/{st.session_state.cleaning_interval}"
        )
    )

    c3.metric(
        "One-run consumption",
        f"{flow * duration:.1f} mL"
    )

    if cleaning_due:

        st.error(
            "CLEANING REQUIRED — "
            "Run simulation blocked until cleaning "
            "is documented on page 2."
        )

    elif (
        st.session_state.stock_ml
        < flow * duration
    ):

        st.warning(
            "Insufficient mobile phase "
            "for the next simulated run."
        )

    if st.button(
        "▶ Simulate one completed LC-MS run",
        disabled=(
            cleaning_due
            or st.session_state.stock_ml
            < flow * duration
        )
    ):

        st.session_state.stock_ml -= (
            flow * duration
        )

        st.session_state.completed_runs += 1

        st.session_state.runs_since_cleaning += 1

        st.session_state.run_logs.append({
            "Timestamp": now(),
            "Instrument": INSTRUMENT_ID,
            "Run number":
                st.session_state.completed_runs,
            "Flow (mL/min)": flow,
            "Duration (min)": duration,
            "Consumed (mL)": round(
                flow * duration, 3
            ),
            "Runs since cleaning":
                st.session_state.runs_since_cleaning,
            "Mode": "SIMULATION"
        })

        st.rerun()


# ============================================================
# PAGE 2 — CLEANING, MAINTENANCE, CALIBRATION
# ============================================================

elif page == "🧼 Reinigung & Wartung":

    page_header(
        "🧼 Cleaning, Maintenance & Calibration",
        "Service Control | Simulation"
    )

    st.info(
        f"Preparation station: {STATION_ID} "
        f"| Virtual LC-MS: {INSTRUMENT_ID}"
    )

    # --------------------------------------------------------
    # CLEANING INTERVAL
    # --------------------------------------------------------

    st.header("1. Cleaning Interval")

    count = st.session_state.runs_since_cleaning
    threshold = st.session_state.cleaning_interval

    c1, c2 = st.columns(2)

    c1.metric(
        "Runs since cleaning",
        f"{count}/{threshold}"
    )

    c2.metric(
        "Runs remaining",
        max(0, threshold - count)
    )

    st.progress(
        min(count / threshold, 1.0)
    )

    if count >= threshold:

        st.error(
            "CLEANING DUE — "
            "Additional simulated runs are blocked."
        )

    else:

        st.success(
            "Cleaning threshold not reached."
        )

    # --------------------------------------------------------
    # CLEANING CHECKLIST
    # --------------------------------------------------------

    st.header("2. Cleaning Checklist")

    clean_checks = {
        "Liquid lines flushed":
            st.checkbox("Liquid lines flushed"),

        "Preparation vessel cleaned":
            st.checkbox("Preparation vessel cleaned"),

        "Waste container checked":
            st.checkbox("Waste container checked"),

        "Visual inspection completed":
            st.checkbox("Visual inspection completed")
    }

    # --------------------------------------------------------
    # MAINTENANCE
    # --------------------------------------------------------

    st.header("3. Maintenance Checklist")

    maintain_checks = {
        "Pump inspected":
            st.checkbox("Pump inspected"),

        "Valves inspected":
            st.checkbox("Valves inspected"),

        "Leak inspection completed":
            st.checkbox("Leak inspection completed")
    }

    # --------------------------------------------------------
    # CALIBRATION
    # --------------------------------------------------------

    st.header("4. Calibration Checklist")

    calibrate_checks = {
        "Balance calibration status verified":
            st.checkbox(
                "Balance calibration status verified"
            ),

        "pH meter calibration status verified":
            st.checkbox(
                "pH meter calibration status verified"
            ),

        "Conductivity meter calibration status verified":
            st.checkbox(
                "Conductivity meter calibration status verified"
            )
    }

    st.caption(
        "Maintenance and calibration are documented "
        "separately and do not need to be repeated "
        "after every cleaning cycle."
    )

    st.divider()

    # --------------------------------------------------------
    # SERVICE DOCUMENTATION
    # --------------------------------------------------------

    st.header("5. Service Documentation")

    c1, c2 = st.columns(2)

    with c1:

        operator = st.selectbox(
            "Responsible operator",
            ["EK", "YB", "IB", "HM"],
            index=None,
            placeholder="Select initials"
        )

    with c2:

        minutes = st.number_input(
            "Cleaning duration (minutes)",
            min_value=0,
            max_value=240,
            value=15
        )

    problem = st.checkbox(
        "Cleaning problem detected"
    )

    notes = st.text_area(
        "Deviations / corrective actions (optional)",
        placeholder=(
            "Only describe deviations "
            "or corrective actions if necessary."
        ),
        height=85
    )

    # Automatically calculate cleaning result
    if problem:

        cleaning_status = "FAILED"

    elif (
        all(clean_checks.values())
        and operator
        and minutes > 0
    ):

        cleaning_status = "PASSED"

    else:

        cleaning_status = "PENDING"

    maintenance_status = (
        "VERIFIED"
        if all(maintain_checks.values())
        else "NOT FULLY VERIFIED"
    )

    calibration_status = (
        "VERIFIED"
        if all(calibrate_checks.values())
        else "NOT FULLY VERIFIED"
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Cleaning result",
        cleaning_status
    )

    c2.metric(
        "Maintenance",
        maintenance_status
    )

    c3.metric(
        "Calibration",
        calibration_status
    )

    st.caption(
        "PASSED means the checklist has been "
        "declared complete. It does not prove "
        "physical cleaning effectiveness."
    )

    if st.button(
        "Save service record",
        type="primary"
    ):

        record = {
            "timestamp": now(),
            "instrument": INSTRUMENT_ID,
            "station": STATION_ID,
            "operator": operator or "NOT SELECTED",
            "cleaning_duration_min": minutes,
            "cleaning_status": cleaning_status,
            "maintenance_status":
                maintenance_status,
            "calibration_status":
                calibration_status,
            "runs_before_cleaning": count,
            "cleaning_checks": clean_checks,
            "maintenance_checks":
                maintain_checks,
            "calibration_checks":
                calibrate_checks,
            "problem_reported": problem,
            "notes": notes,
            "mode": (
                "SIMULATION / MANUAL DECLARATION"
            )
        }

        st.session_state.service_logs.append(
            record
        )

        if cleaning_status == "PASSED":

            st.session_state.runs_since_cleaning = 0

            st.success(
                "Simulated cleaning recorded. "
                "Run counter reset to 0."
            )

        else:

            st.warning(
                "Service record saved, "
                "but cleaning is not complete. "
                "Run counter unchanged."
            )

        st.rerun()


# ============================================================
# PAGE 3 — DIGITAL DOCUMENTATION
# ============================================================

elif page == "📋 Dokumentation":

    page_header(
        "📋 Digital Laboratory Documentation",
        "Batch Records | Materials | Runs | Service"
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Simulated runs",
        st.session_state.completed_runs
    )

    c2.metric(
        "Recorded batches",
        len(st.session_state.batches)
    )

    c3.metric(
        "Service entries",
        len(st.session_state.service_logs)
    )

    # --------------------------------------------------------
    # RUN HISTORY
    # --------------------------------------------------------

    st.header("1. Run History")

    if st.session_state.run_logs:

        st.dataframe(
            pd.DataFrame(
                st.session_state.run_logs
            ),
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "No simulated runs recorded yet."
        )

    # --------------------------------------------------------
    # BATCH HISTORY
    # --------------------------------------------------------

    st.header("2. Batch History & QC")

    if st.session_state.batches:

        st.dataframe(
            pd.DataFrame([
                flatten_batch(b)
                for b in st.session_state.batches
            ]),
            use_container_width=True,
            hide_index=True
        )

        selected_id = st.selectbox(
            "Select batch",
            [
                b["batch_id"]
                for b in st.session_state.batches
            ]
        )

        selected = next(
            b
            for b in st.session_state.batches
            if b["batch_id"] == selected_id
        )

        with st.expander(
            "Complete batch record, "
            "material lots and QC results",
            expanded=True
        ):

            st.json(selected)

    else:

        st.info(
            "No simulated batches recorded yet."
        )

    # --------------------------------------------------------
    # SERVICE HISTORY
    # --------------------------------------------------------

    st.header(
        "3. Cleaning, Maintenance & Calibration History"
    )

    if st.session_state.service_logs:

        service_table = pd.DataFrame([
            {
                "Date": r["timestamp"],
                "Operator": r["operator"],
                "Cleaning": r["cleaning_status"],
                "Maintenance": r["maintenance_status"],
                "Calibration": r["calibration_status"],
                "Duration (min)":
                    r["cleaning_duration_min"]
            }
            for r in st.session_state.service_logs
        ])

        st.dataframe(
            service_table,
            use_container_width=True,
            hide_index=True
        )

        with st.expander(
            "Full service records"
        ):

            st.json(
                st.session_state.service_logs
            )

    else:

        st.info(
            "No service entries recorded yet."
        )

    # --------------------------------------------------------
    # EXPORT
    # --------------------------------------------------------

    st.header("4. Export Documentation")

    c1, c2, c3 = st.columns(3)

    with c1:

        st.download_button(
            "⬇ Download Excel (.xlsx)",
            data=workbook_bytes(),
            file_name=(
                "LabStream_simulation_records.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )

    with c2:

        st.download_button(
            "⬇ Download PDF (.pdf)",
            data=pdf_bytes(),
            file_name=(
                "LabStream_simulation_report.pdf"
            ),
            mime="application/pdf"
        )

    with c3:

        complete = {
            "mode": "SIMULATION",
            "instrument": INSTRUMENT_ID,
            "station": STATION_ID,
            "batches": st.session_state.batches,
            "runs": st.session_state.run_logs,
            "service": st.session_state.service_logs
        }

        st.download_button(
            "⬇ Download JSON",
            data=json.dumps(
                complete,
                indent=2,
                ensure_ascii=False
            ),
            file_name=(
                "LabStream_simulation_records.json"
            ),
            mime="application/json"
        )

    st.warning(
        "Records are stored only during the "
        "current Streamlit session. "
        "Download the documentation before "
        "ending the session. "
        "This prototype is not a validated "
        "GMP-compliant laboratory system."
    )
