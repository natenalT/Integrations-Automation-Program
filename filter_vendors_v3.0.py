# =========================================================
# filter_vendors_v3.0.py
# Hybrid version with Matt’s logic + automatic COCID/HCIS
# =========================================================

import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# =====================
# 📁 File Paths
# =====================
unfiltered_folder = r'C:\Users\joc4126\Desktop\CSV_File\Unfiltered'
filtered_folder = r'C:\Users\joc4126\Desktop\CSV_File\Filtered'
crosswalk_path = r'C:\Users\joc4126\Desktop\CSV_File\Files_To_Import\Data_Type_Crosswalk_Final.xlsx'
connection_path = r'C:\Users\joc4126\Desktop\CSV_File\Files_To_Import\Connection_Type_Crosswalk.xlsx'

os.makedirs(filtered_folder, exist_ok=True)

# =====================
# 📖 Read reference files
# =====================
crosswalk_df = pd.read_excel(crosswalk_path, usecols="A:F")
connection_df = pd.read_excel(connection_path, usecols="A:C")

crosswalk_df.columns = crosswalk_df.columns.str.strip()
connection_df.columns = connection_df.columns.str.strip()

# Mapping dictionaries
message_type_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 4]))  # Col E
data_type_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 5]))     # Col F
expanse_mnemonic_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 2]))  # Col C

connection_type_map = dict(zip(connection_df.iloc[:, 0], connection_df.iloc[:, 1]))  # Col A->B
contract_map = dict(zip(connection_df.iloc[:, 0], connection_df.iloc[:, 2]))         # Col A->C

wave_iteration = "Wave 2D"
division = "Gulf Coast Division"
default_hcis = "Meditech Expanse"

# =====================
# 🧠 Helper Functions
# =====================
def get_data_flow(row):
    ib = str(row.get('ib_product_name', ''))
    ob = str(row.get('ob_product_name', ''))

    if 'HCA 5.6' in ob or 'MPF-Patient Folder/HPF Facility' in ob:
        return "Inbound"
    elif 'HCA 5.6' in ib:
        return "Outbound"
    return "Inbound"

def get_product(row):
    direction = get_data_flow(row)
    return row.get('ob_vendor_name', '') if direction == 'Outbound' else row.get('ib_vendor_name', '')

def get_interface(row):
    data_type = data_type_map.get(row.get('ib_datatype', ''), '')
    direction = get_data_flow(row)
    vendor_name = row.get('ob_vendor_name', '') if direction == 'Outbound' else row.get('ib_vendor_name', '')
    product = row.get('Vendor', '')
    return f"{vendor_name} ({product}) - {data_type}"

def get_connection_type(row):
    ob = str(row.get('ob_product_name', '')).lower()
    ib = str(row.get('ib_product_name', '')).lower()
    combined = f"{ob} {ib}"
    if "kafka" in combined:
        return "Kafka"
    elif "waterpark" in combined:
        return "Waterpark"
    return "Cloverleaf"

def get_facility_type(facility_name):
    if "Imaging" in facility_name:
        return "Imaging Center"
    elif "Oncology" in facility_name or "Cancer" in facility_name:
        return "Oncology Clinic"
    return "Acute"

def clean_facility_name(filename):
    base = filename.replace('.csv', '')
    for prefix in ['COCAT_', 'COCCL_', 'COCHRA_', 'COCKN_', 'COCMHB_', 'COCNW_', 'COCQR_', 'COCUH_', 'COCUHA_', 'COCWS_', 'COCXG_']:
        if base.startswith(prefix):
            base = base[len(prefix):]
    base = base.replace('-InterfaceMigrationReport', '')
    return base.replace('_', ' ').replace('-', ' ').strip()

def extract_cocid_from_filename(filename):
    return os.path.basename(filename).split('_')[0]

# =====================
# 🧩 Main Processing Function
# =====================
def process_facility(input_csv_path):
    filename = os.path.basename(input_csv_path)
    facility_name = clean_facility_name(filename)
    cocid = extract_cocid_from_filename(filename)
    print(f"Processing {facility_name}...")

    df = pd.read_csv(input_csv_path)
    possible_cols = ['ob_product_name', 'ib_product_name', 'Vendor', 'ib_vendor_name', 'ob_vendor_name', 'ib_datatype', 'ob_datatype']
    cols_for_dedup = [c for c in possible_cols if c in df.columns]
    if cols_for_dedup:
        df[cols_for_dedup] = df[cols_for_dedup].astype(str)
        df = df.drop_duplicates(subset=cols_for_dedup)

    filtered_df = pd.DataFrame({
        'Iteration Path': f"Integration-Tracker\\{wave_iteration}",
        'Area Path': f"Integration-Tracker\\{division}",
        'Work Item Type': 'Interface',
        'Contract Ownership': '',
        'Product Owner': '',
        'Title 1': df.apply(get_product, axis=1),
        'Title 2': df.apply(get_interface, axis=1),
        'State': "Contracting in Progress",
        'Facility': facility_name,
        'Facility Type': get_facility_type(facility_name),
        'COCID': cocid,
        'HCIS': default_hcis,
        'Division': division,
        'Product': df.apply(get_product, axis=1),
        'Connection Type': df.apply(get_connection_type, axis=1),
        'Data Flow Direction': df.apply(get_data_flow, axis=1),
        'Expanse Interface Mnemonic': df['ib_datatype'].map(expanse_mnemonic_map),
        'Message Type': df['ib_datatype'].map(message_type_map),
        'Data Type': df['ib_datatype'].map(data_type_map)
    }).dropna(subset=['Title 1', 'Title 2'])

    filtered_df = filtered_df.sort_values(by=['Title 1', 'Title 2'], ascending=[True, True])

    output_rows = []
    for product, group in filtered_df.groupby(['Title 1', 'Title 2'], sort=False):
        vendor_row = {
            'Iteration Path': f"Integration-Tracker\\{wave_iteration}",
            'Area Path': f"Integration-Tracker\\{division}",
            'Work Item Type': 'Product',
            'Contract Ownership': '',
            'Product Owner': '',
            'Title 1': product,
            'Title 2': '',
            'State': "1. Contracting In Progress",
            'Facility': facility_name,
            'Facility Type': get_facility_type(facility_name),
            'COCID': cocid,
            'HCIS': default_hcis,
            'Division': division,
            'Product': group['Product'].iloc[0],
            'Connection Type': group['Connection Type'].iloc[0],
            'Data Flow Direction': '',
            'Expanse Interface Mnemonic': '',
            'Message Type': '',
            'Data Type': ''
        }
        output_rows.append(vendor_row)
        for _, row in group.iterrows():
            row_copy = row.copy()
            row_copy['Title 1'] = ''
            output_rows.append(row_copy)

    final_df = pd.DataFrame(output_rows)

    output_excel_path = os.path.join(filtered_folder, f"{facility_name.replace(' ', '_')}.xlsx")
    final_df.to_excel(output_excel_path, index=False)

    wb = load_workbook(output_excel_path)
    ws = wb.active
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center')

    for row in range(2, ws.max_row + 1):
        if not ws.cell(row, 2).value:
            for col in range(1, ws.max_column + 1):
                ws.cell(row, col).font = bold_font
                ws.cell(row, col).alignment = center_align

    wb.save(output_excel_path)
    print(f"✅ {facility_name} exported successfully.")

# =====================
# 🚀 Main Runner
# =====================
def main():
    csv_files = [os.path.join(unfiltered_folder, f) for f in os.listdir(unfiltered_folder) if f.lower().endswith('.csv')]
    if not csv_files:
        print(f"No CSV files found in {unfiltered_folder}")
        return

    for csv_path in csv_files:
        try:
            process_facility(csv_path)
        except Exception as e:
            print(f"❌ Error processing {csv_path}: {e}")

if __name__ == "__main__":
    main()
