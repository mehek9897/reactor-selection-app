import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment


def load_reactor_data(filepath):
    df = pd.read_excel(filepath)
    df.columns = df.columns.str.strip().str.lower()

    rename_map = {
        "vessel id": "reactor id",
        "min sensing volume": "min sensing",
        "min stirring volume": "min stirring",
        "capacity": "max volume",
        "moc": "moc",
        "utilities": "utilities",
        "agitator": "agitator"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    if "utilities" not in df.columns:
        st.error("❌ 'utilities' column not found in the uploaded Excel. Please ensure it's named correctly.")
        return pd.DataFrame()

    df["moc"] = df["moc"].str.upper().replace({"ALL GLASS": "GLR"})
    df["materials"] = df["moc"].apply(lambda x: [m.strip() for m in x.split("/")])
    df["thermal options"] = df["utilities"].astype(str).apply(lambda x: [t.strip().upper() for t in x.split(",")])
    df["agitator"] = df["agitator"].astype(str).str.upper()
    return df[["reactor id", "min sensing", "min stirring", "max volume", "materials", "thermal options", "agitator"]]


def collect_unit_operation(unit_op_id):
    steps = []
    total_volume = 0
    first_step_volume = None

    st.subheader("🧪 Add Steps to Unit Operation")
    step_count = 0
    add_more_steps = True

    while add_more_steps:
        step_count += 1
        st.markdown(f"### Step {step_count}")
        operation = st.selectbox(f"Select operation type for Step {step_count}", ["charge", "addition"], key=f"op_{unit_op_id}_{step_count}")
        material = st.selectbox(f"Select material type for Step {step_count}", ["reagent 1", "reagent 2", "reagent 3", "ksm", "solvent"], key=f"mat_{unit_op_id}_{step_count}")
        volume = st.number_input(f"Enter volume (L) for Step {step_count}", min_value=0.0, key=f"vol_{unit_op_id}_{step_count}")

        actual_volume = volume
        if material == "ksm":
            percentage = st.number_input(f"Enter percentage of KSM for Step {step_count}", min_value=0.0, max_value=100.0, key=f"ksm_{unit_op_id}_{step_count}")
            if percentage > 0:
                actual_volume = volume / (percentage / 100)

        if first_step_volume is None:
            first_step_volume = actual_volume

        total_volume += actual_volume
        steps.append({
            "unit_op": unit_op_id,
            "step": step_count,
            "operation": operation,
            "material": material,
            "input_volume": volume,
            "actual_volume": actual_volume,
            "accumulated_volume": total_volume
        })

        add_more_steps = st.radio(f"Add another step after Step {step_count}?", ["yes", "no"], index=1, key=f"cont_{unit_op_id}_{step_count}") == "yes"

    return first_step_volume, total_volume, steps


def filter_reactors(df, user_input, first_step_vol, total_vol):
    df = df[(df["min sensing"] <= first_step_vol) & (df["min stirring"] <= first_step_vol)]
    vol_limit = 0.7 if user_input["pressurized"] == "yes" else 0.95
    df = df[df["max volume"] * vol_limit >= total_vol]

    if user_input["ph_condition"] == "basic":
        allowed = ["SSR", "HAR"]
    elif user_input["ph_condition"] == "acidic":
        allowed = ["GLR", "HAR"]
    elif user_input["ph_condition"] == "neutral":
        allowed = ["GLR", "SSR", "HAR"]
    elif user_input["ph_condition"] == "coupon":
        mat = user_input["coupon_materials"][0].strip().upper()
        if user_input["corrosion_rate"] < 0.1:
            allowed = [mat]
        else:
            st.error("❌ Corrosion rate too high for this material.")
            return pd.DataFrame()
    else:
        allowed = []

    df = df[df["materials"].apply(lambda mats: any(m in mats for m in allowed))]

    temp = user_input["temperature"]
    if 10 <= temp <= 20:
        thermal = ["CHB"]
    elif 20 < temp <= 35:
        thermal = ["CT"]
    elif 20 < temp <= 90:
        thermal = ["HW"]
    else:
        thermal = ["LPS", "HOT OIL", "EJECTION CONDENSATE"]

    df = df[df["thermal options"].apply(lambda opts: any(t in opts for t in thermal))]

    preferred = []
    if user_input["reaction_nature"] == "homogeneous":
        preferred = ["PROPELLOR", "PBT", "RCI", "ANCHOR", "CBRT"]
    elif user_input["reaction_nature"] == "heterogeneous":
        if user_input["reaction_subtype"] == "biphasic":
            preferred = ["PROPELLOR", "PBT", "CBRT", "RCI"]
        elif user_input["reaction_subtype"] == "solid-liquid":
            preferred = ["PROPELLOR", "PBT", "CBRT", "RCI", "ANCHOR"]
        elif user_input["reaction_subtype"] == "gas-liquid":
            preferred = ["RUSTON", "DISC"]

    df["Preference Match"] = df["agitator"].apply(lambda a: "✅" if any(p in a for p in preferred) else "⚠️")
    return df


def export_steps_to_excel(steps_by_unitop):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        all_steps = []
        for unitop_id, (steps, selected_reactor) in enumerate(steps_by_unitop, start=1):
            for s in steps:
                all_steps.append({
                    "Unit Operation": unitop_id,
                    "Operation": s["operation"],
                    "Material": s["material"],
                    "Volume Added (L)": s["actual_volume"],
                    "Accumulated Volume (L)": s["accumulated_volume"],
                    "Reactor ID": selected_reactor
                })

        df_export = pd.DataFrame(all_steps)
        df_export.to_excel(writer, index=False, sheet_name="Steps", startrow=2)

    
        ws = writer.sheets["Steps"]

        for col in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max_length + 2

        current_row = 3
        prev_op = None
        for i, row in enumerate(df_export.itertuples(index=False), start=current_row):
            if prev_op != row[0]:
                start_row = i
                count = df_export["Unit Operation"].tolist().count(row[0])
                ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + count - 1, end_column=1)
                cell = ws.cell(row=start_row, column=1)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                prev_op = row[0]

    buffer.seek(0)
    return buffer


def main():
    st.set_page_config("Reactor Selector", layout="centered")
    st.title("🧪 Reactor Selection Web App")
    uploaded_file = st.file_uploader("📄 Upload reactor database", type="xlsx")

    if uploaded_file:
        df = load_reactor_data(uploaded_file)
        if df.empty:
            return

        all_results = []
        step_tracking = []

        st.header("🔧 Enter Process Conditions")
        batch_id = 0
        while True:
            batch_id += 1
            st.markdown(f"## 🧾 Unit Operation {batch_id}")

            pressurized = st.radio("1. Is the reaction pressurized?", ["yes", "no"], key=f"pres_{batch_id}")
            ph_condition = st.selectbox("2. pH condition", ["basic", "acidic", "neutral", "coupon"], key=f"ph_{batch_id}")
            corrosion_rate = 0
            coupon_materials = []

            if ph_condition == "coupon":
                corrosion_rate = st.number_input("Corrosion rate (mm/year)", min_value=0.0, key=f"cr_{batch_id}")
                coupon_materials = [st.text_input("Material for coupon study", key=f"mat_{batch_id}").upper()]

            temperature = st.number_input("3. Process temperature (°C)", min_value=0.0, key=f"temp_{batch_id}")
            reaction_nature = st.selectbox("4. Nature of reaction", ["none", "homogeneous", "heterogeneous"], key=f"rn_{batch_id}")
            reaction_subtype = None
            if reaction_nature == "heterogeneous":
                reaction_subtype = st.selectbox("Subtype", ["biphasic", "solid-liquid", "gas-liquid"], key=f"rs_{batch_id}")

            st.markdown("---")
            first_vol, total_vol, step_log = collect_unit_operation(batch_id)

            if st.button(f"🔍 Submit Unit Operation {batch_id}", key=f"submit_{batch_id}"):
                user_input = {
                    "pressurized": pressurized,
                    "ph_condition": ph_condition,
                    "corrosion_rate": corrosion_rate,
                    "coupon_materials": coupon_materials,
                    "temperature": temperature,
                    "reaction_nature": reaction_nature,
                    "reaction_subtype": reaction_subtype
                }

                matched_df = filter_reactors(df.copy(), user_input, first_vol, total_vol)

                if not matched_df.empty:
                    styled = matched_df[["reactor id", "min sensing", "min stirring", "max volume", "agitator", "Preference Match"]]
                    st.success(f"✅ Reactors matching Unit Operation {batch_id}")
                    selected_reactor = st.selectbox("Select one reactor to use:", styled["reactor id"].tolist(), key=f"sel_reactor_{batch_id}")
                    st.dataframe(styled.style.applymap(
                        lambda v: "background-color: #d4edda" if v == "✅" else "background-color: #fff3cd",
                        subset=["Preference Match"]
                    ))
                    all_results.append(styled)
                    step_tracking.append((step_log, selected_reactor))
                else:
                    st.warning("⚠️ No matching reactors found for this unit operation.")

            another = st.radio(f"Add another Unit Operation?", ["no", "yes"], index=0, key=f"another_{batch_id}")
            if another == "no":
                break

        if step_tracking:
            excel_buffer = export_steps_to_excel(step_tracking)
            st.download_button(
                "📥 Download Steps Summary",
                data=excel_buffer.getvalue(),
                file_name="unit_op_steps.xlsx"
            )

if __name__ == "__main__":
    main()