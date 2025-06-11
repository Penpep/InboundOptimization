import streamlit as st
st.set_page_config(page_title="Inbound Dock Inventory Analysis", layout='wide',
    page_icon='smiley')


import pandas as pd
import math
from datetime import datetime, timedelta
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import io
import tempfile


def generate_times(start, end, cadence):
    interval = (end - start) / cadence
    return [start + i * interval for i in range(cadence)]

def generate_deliveries(qty, pack_size, cadence, shft_hrs, cons_rate):
    deliveries = []
    interval = shft_hrs / cadence
    total_units_delivered = 0
    total_units_needed = qty

    for i in range(cadence):
        units_needed = cons_rate * interval
        remaining_units = total_units_needed - total_units_delivered
        units_to_deliver = min(math.ceil(units_needed), remaining_units)
        packages = math.ceil(units_to_deliver / pack_size)
        deliveries.append(packages)
        total_units_delivered += packages * pack_size
        if total_units_delivered >= total_units_needed:
            break
    while len(deliveries) < cadence:
        deliveries.append(0)
    return deliveries

# === Space Analysis =======
def get_dock_inventory_peaks_per_part(deliveries, pack_size, consumption_rate, shift_hours, max_lineside, min_lineside):
    if not deliveries or pack_size <= 0 or consumption_rate <= 0:
        return [0] * len(deliveries)

    interval = shift_hours / len(deliveries)
    dock_inventory_units = 0
    lineside_inventory_units = 0
    dock_timeline = []

    for i, delivery in enumerate(deliveries):
        delivered_units = delivery * pack_size
        dock_inventory_units += delivered_units

        # === Record peak BEFORE lineside pull ===
        dock_timeline.append(dock_inventory_units)

        consumption = consumption_rate * interval
        pull_needed = max(consumption - lineside_inventory_units, 0)
        actual_pull = min(pull_needed, dock_inventory_units)

        # Pull to lineside
        dock_inventory_units -= actual_pull
        lineside_inventory_units += actual_pull

        # 4. Consume from lineside
        lineside_inventory_units = max(0, lineside_inventory_units - consumption)

    return dock_timeline

def run_analysis(uploaded_file):
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb['Inbound']

    cadence_shift_1 = ws['B5'].value
    cadence_shift_2 = ws['B13'].value
    rate_per_line = ws['B6'].value
    lines_shift_1 = ws['B9'].value
    lines_shift_2 = ws['B10'].value
    pallet_dock_space = ws['B2'].value
    box_dock_space = ws['B3'].value
    total_line_side = ws['F6'].value
    pallets_per_line = ws['F7'].value

    start_shift_1 = datetime.strptime("6:15", "%H:%M")
    end_shift_1 = datetime.strptime("15:00", "%H:%M")
    start_shift_2 = datetime.strptime("15:00", "%H:%M")
    end_shift_2 = datetime.strptime("23:15", "%H:%M")
    shift_1_hours = (end_shift_1 - start_shift_1).seconds / 3600
    shift_2_hours = (end_shift_2 - start_shift_2).seconds / 3600

    time_1 = generate_times(start_shift_1, end_shift_1, cadence_shift_1)
    time_2 = generate_times(start_shift_2, end_shift_2, cadence_shift_2)

    df_bom = pd.read_excel(uploaded_file, sheet_name='Inbound', skiprows=15, header=None)
    df_bom.columns = [
        'Part Number', 'Description', 'Quantity / Unit', 'Needed per day',
        'Quantity Needed for Shift 1', 'Quantity Needed for Shift 2',
        'Pallets Utilized for Shift 1', 'Pallets Utilized for Shift 2',
        'Pallets Utilized per day', 'Replenishment Frequency',
        'Max Quantity during Arrival', 'Arrival Rate per X hours',
        'Departure rate Pallets per hour',
        'Consumption Rate Units/ Hour Shift 1',
        'Consumption Rate / Hour Shift 2', 'Standard Pack Size', 'Package Type',
        'Boxes to Pallets shift 1', 'Maximum Storage on Lineside',
        'Minimum Storage on Lineside', 'Helper Column 1', 'Helper Column 2',
        'Boxes to Pallets shift 2']

    columns = ['Part Number', 'Package Type']
    columns += [f"Delivery {i+1} (S1 - {t.strftime('%I:%M %p')})" for i, t in enumerate(time_1)]
    columns += [f"Delivery {i+1} (S2 - {t.strftime('%I:%M %p')})" for i, t in enumerate(time_2)]

    delivery_plan = []

    for i, row in df_bom.iterrows():
        part = row['Part Number']
        pkg_type = row['Package Type']
        pack_size = row['Standard Pack Size']
        qty1 = row['Quantity Needed for Shift 1']
        qty2 = row['Quantity Needed for Shift 2']
        cons_1 = row['Consumption Rate Units/ Hour Shift 1']
        cons_2 = row['Consumption Rate / Hour Shift 2']

        deliveries_1 = generate_deliveries(qty1, pack_size, cadence_shift_1, shift_1_hours, cons_1)
        deliveries_2 = generate_deliveries(qty2, pack_size, cadence_shift_2, shift_2_hours, cons_2)

        delivery_plan.append([part, pkg_type] + deliveries_1 + deliveries_2)

    df_output = pd.DataFrame(delivery_plan, columns=columns)

    df_boxes = df_output[df_output['Package Type'] == 'Box']
    df_pallets = df_output[df_output['Package Type'] == 'Pallet']

    box_total = ['TOTAL - BOX', 'Box'] + list(df_boxes.iloc[:, 2:].sum())
    pallet_total = ['TOTAL - PALLET', 'Pallet'] + list(df_pallets.iloc[:, 2:].sum())
    box_sums = df_boxes.iloc[:, 2:].sum()
    pallet_sums = df_pallets.iloc[:, 2:].sum()

    # Compute space utilization 
    utilization_pallet = ['Space Utilization Pallet', '']
    utilization_box = ['Space Utilization Box', '']
    total_lanes_needed = ['Lanes Utilized', '']
    percent_lanes = ['Percent of Lanes Utilized', '']
    for box, pallet in zip(box_sums, pallet_sums):
        percent_box = (box / box_dock_space) * 100 if box_dock_space else 0
        percent_pallet = (pallet / pallet_dock_space) * 100 if pallet_dock_space else 0

        lanes_used = math.ceil(pallet/ pallets_per_line) 
        lanes_percent = 100*lanes_used / total_line_side 

        utilization_pallet.append(round(percent_pallet, 1))
        utilization_box.append(round(percent_box, 1))

        total_lanes_needed.append(lanes_used)
        percent_lanes.append(round(lanes_percent,1))

    df_output.loc[len(df_output)] = box_total
    df_output.loc[len(df_output)] = pallet_total
    df_output.loc[len(df_output)] = utilization_box
    df_output.loc[len(df_output)] = utilization_pallet
    df_output.loc[len(df_output)] = total_lanes_needed
    df_output.loc[len(df_output)] = percent_lanes

    #=============================================================================================

    space_records = []

    for idx, row in df_bom.iterrows():
        part = row['Part Number']
        pack_size = row['Standard Pack Size']
        pkg_type = row['Package Type']
        cons_1 = row['Consumption Rate Units/ Hour Shift 1'] * 0.9
        cons_2 = row['Consumption Rate / Hour Shift 2'] * 0.9

        if pd.isna(pack_size) or pack_size <= 0:
            continue

        deliveries_1 = df_output.loc[idx, df_output.columns.str.contains(r"S1 -")].tolist()
        deliveries_2 = df_output.loc[idx, df_output.columns.str.contains(r"S2 -")].tolist()

        max_lineside_1 = lines_shift_1 * row['Maximum Storage on Lineside']
        min_lineside_1 = lines_shift_1 * row['Minimum Storage on Lineside']
        max_lineside_2 = lines_shift_2 * row['Maximum Storage on Lineside']
        min_lineside_2 = lines_shift_2 * row['Minimum Storage on Lineside']

        timeline_1 = get_dock_inventory_peaks_per_part(
            deliveries_1, pack_size, cons_1, shift_1_hours, max_lineside_1, min_lineside_1
        )
        timeline_2 = get_dock_inventory_peaks_per_part(
            deliveries_2, pack_size, cons_2, shift_2_hours, max_lineside_2, min_lineside_2
        )

        for i, inv_units in enumerate(timeline_1):
            space_records.append({
                'Part Number': part,
                'Shift': 1,
                'Delivery Label': f"Delivery {i+1} (S1 - {time_1[i].strftime('%I:%M %p')})",
                'Inventory Packages': inv_units // pack_size,
                'Package Type': pkg_type
            })

        for i, inv_units in enumerate(timeline_2):
            space_records.append({
                'Part Number': part,
                'Shift': 2,
                'Delivery Label': f"Delivery {i+1} (S2 - {time_2[i].strftime('%I:%M %p')})",
                'Inventory Packages': inv_units // pack_size,
                'Package Type': pkg_type
            })

    df_space = pd.DataFrame(space_records)

    # === Format flat dock inventory table ===
    flat_records = []
    part_order = df_bom['Part Number'].tolist()

    for part in part_order:
        part_rows = df_space[df_space['Part Number'] == part]
        if part_rows.empty:
            continue
        pkg_type = part_rows['Package Type'].iloc[0]
        row = {'Part #': part, 'TYPE': pkg_type}
        for _, rec in part_rows.iterrows():
            row[rec['Delivery Label']] = rec['Inventory Packages']
        flat_records.append(row)

    df_dock_space = pd.DataFrame(flat_records).fillna(0)

    # === Totals and utilization ===
    delivery_headers = [col for col in df_dock_space.columns if col.startswith('Delivery')]
    df_boxes = df_dock_space[df_dock_space['TYPE'] == 'Box']
    df_pallets = df_dock_space[df_dock_space['TYPE'] == 'Pallet']

    row_total_box = ['Total BOXES', ''] + list(df_boxes[delivery_headers].sum().values)
    row_total_pallet = ['Total PALLETS', ''] + list(df_pallets[delivery_headers].sum().values)
    row_util_pallet = ['Percent Utilization Pallet', '']
    row_util_box = ['Percent Utilization Box', '']
    total_lanes_needed = ['Lanes Utilized', '']
    percent_lanes = ['Percent of Lanes Utilized', '']

    box_sums = df_boxes[delivery_headers].sum()
    pallet_sums = df_pallets[delivery_headers].sum()

    for col in delivery_headers:
        box_count = box_sums[col]
        pallet_count = pallet_sums[col]

        percent_pallet = (pallet_count / pallet_dock_space) * 100 if pallet_dock_space else 0
        percent_box = (box_count / box_dock_space) * 100 if box_dock_space else 0
        row_util_pallet.append(round(percent_pallet, 1))
        row_util_box.append(round(percent_box, 1))
        
        lanes_used = math.ceil(pallet_count / pallets_per_line)
        lanes_percent = 100 * lanes_used / total_line_side
        total_lanes_needed.append(lanes_used)
        percent_lanes.append(round(lanes_percent, 1))

    df_dock_space.loc[len(df_dock_space)] = row_total_box
    df_dock_space.loc[len(df_dock_space)] = row_total_pallet
    df_dock_space.loc[len(df_dock_space)] = row_util_box
    df_dock_space.loc[len(df_dock_space)] = row_util_pallet
    df_dock_space.loc[len(df_dock_space)] = total_lanes_needed
    df_dock_space.loc[len(df_dock_space)] = percent_lanes
    return df_output, df_dock_space

# Streamlit interface
st.title("Inbound Delivery Planning Tool")

uploaded_file = st.file_uploader("Upload your BOM Excel file", type=["xlsx", "xlsm"])
if uploaded_file:
    st.success("File uploaded successfully.")
    if st.button("Generate Delivery Plan"):
        with st.spinner("Processing..."):
            df_output, df_dock_space = run_analysis(uploaded_file)

        st.subheader("Delivery Plan Output")

        st.dataframe(df_output)
        st.subheader("Dock Inventory Space Per Part (Pallet Equivalents)")
        st.dataframe(df_dock_space)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_output.to_excel(writer, sheet_name="Delivery Plan", index=False)
            df_dock_space.to_excel(writer, sheet_name="Dock Inventory Space", index=False)
        st.download_button("Download Excel", output.getvalue(), file_name="Delivery_Plan.xlsx")
        
