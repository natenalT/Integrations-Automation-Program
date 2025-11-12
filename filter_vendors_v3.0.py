# =========================================================
# filter_vendors_v3.0.py
# Final Combined Version (Matt + Nate)
# =========================================================

import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# =====================
# 📁 File Paths
# =====================
unfiltered_folder = r'C:\Users\joc4126\Desktop\CSV_File\Unfiltered'
filtered_folder = r'C:\Users\joc4126\Desktop\CSV_File\Filtered_V3.0'
crosswalk_path = r'C:\Users\joc4126\Desktop\CSV_File\Files_To_Import\Data_Type_Crosswalk_Final.xlsx'
connection_path = r'C:\Users\joc4126\Desktop\CSV_File\Files_To_Import\Connection_Type_Crosswalk.xlsx'

os.makedirs(filtered_folder, exist_ok=True)

# =====================
# 📖 Load Reference Files
# =====================

crosswalk_df = pd.read_excel(crosswalk_path, usecols="A:F")
connection_df = pd.read_excel(connection_path, usecols="A:C")

# Remove spaces in column names
crosswalk_df.columns = crosswalk_df.columns.str.strip()
connection_df.columns = connection_df.columns.str.strip()

# Clean and adjust the crosswalk file based on Matt’s notes
# 1. Remove "_ipl" suffix from datatype column (Column A)
# 2. Remove rows 19 and 21 (0-based index 18 and 20)
crosswalk_df.iloc[:, 0] = crosswalk_df.iloc[:, 0].str.replace('_ipl', '', regex=False)
crosswalk_df = crosswalk_df.drop(index=[18, 20], errors='ignore')

# Create mapping dictionaries
message_type_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 4]))  # Column E
data_type_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 5]))     # Column F
expanse_mnemonic_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 2]))  # Column C

connection_type_map = dict(zip(connection_df.iloc[:, 0], connection_df.iloc[:, 1]))  # Col A->B
contract_map = dict(zip(connection_df.iloc[:, 0], connection_df.iloc[:, 2]))         # Col A->C

# =====================
# 📍 Constants
# =====================
wave_iteration = "Wave 2D"
division = "Gulf Coast Division"
default_hcis = "Meditech Expanse"

# =====================
# 🧠 Helper Functions
# =====================

def get_data_flow(row):
    """
    Determines if an interface is inbound or outbound.
    Uses product name fields to decide direction.
    """
    ib = str(row.get('ib_product_name', '')).lower()
    ob = str(row.get('ob_product_name', '')).lower()

    if 'hca 5.6' in ob or 'hpf facility' in ob:
        return "Inbound"
    elif 'hca 5.6' in ib or 'ipeople' in ib:
        return "Outbound"
    return "CHECK THIS ROW"

def get_special_ipeople_values(row):
    """
    Handles special rules for iPeople interfaces.
    If ib_product_name contains 'iPeople' and the datatype/message type
    is 'oeord' or 'sch', then override their values with iPeople-specific ones.
    """
    ib_name = str(row.get('ib_product_name', '')).lower()
    datatype = str(row.get('ib_datatype', '')).lower()
    messagetype = str(row.get('ob_datatype', '')).lower()

    if 'ipeople' in ib_name:
        if datatype == 'oeord' or messagetype == 'oeord':
            return {
                'DataType': 'ORM - iPeople',
                'MessageType': 'iPeople OM Orders'
            }
        elif datatype == 'sch' or messagetype == 'sch':
            return {
                'DataType': 'SIU',
                'MessageType': 'iPeople Scheduling'
            }

    # Default mapping when not iPeople-related
    return {
        'DataType': data_type_map.get(row.get('ib_datatype', ''), ''),
        'MessageType': message_type_map.get(row.get('ib_datatype', ''), '')
    }

def get_product(row):
    """
    Combines vendor and product name for easier grouping.
    If outbound, use ob_vendor_name; otherwise use ib_vendor_name.
    """
    direction = get_data_flow(row)
    product_name = str(row.get('Vendor', '')).strip()

    if direction == 'Outbound':
        vendor = str(row.get('ob_vendor_name', '')).strip()
    else:
        vendor = str(row.get('ib_vendor_name', '')).strip()

    if not vendor:
        vendor = product_name

    return f"{vendor} ({product_name})"

def get_interface(row):
    """
    Builds a combined interface name using the vendor/product and data type.
    """
    data_type = data_type_map.get(row.get('ib_datatype', ''), '')
    vendor_product = get_product(row)
    return f"{vendor_product}-{data_type}"

def get_connection_type(row):
    """
    Looks up the connection type from the connection crosswalk.
    Defaults to Cloverleaf if no match found.
    """
    product = row.get('Vendor', "")
    connection_type = connection_type_map.get(product, "")
    if not connection_type:
        connection_type = "Cloverleaf"
    return connection_type


def get_facility_type(facility_name):
    """
    Classifies facilities by name pattern.
    """
    if "Imaging" in facility_name:
        return "Imaging Center"
    if "Oncology" in facility_name or "Cancer" in facility_name:
        return "Oncology Clinic"
    return "Acute"

def clean_facility_name(filename):
    """
    Removes prefixes and special suffixes from filenames to extract facility names.
    """
    base = filename.replace('.csv', '')
    prefixes = ['COCAT_', 'COCCL_', 'COCHRA_', 'COCKN_', 'COCMHB_', 'COCNW_', 'COCQR_', 'COCUH_', 'COCUHA_', 'COCWS_', 'COCXG_']
    for prefix in prefixes:
        if base.startswith(prefix):
            base = base[len(prefix):]
    base = base.replace('-InterfaceMigrationReport', '')
    return base.replace('_', ' ').replace('-', ' ').strip()

def extract_cocid_from_filename(filename):
    """
    Extracts the COCID (facility code) from the start of the file name.
    """
    return os.path.basename(filename).split('_')[0]

# =====================
# 🧩 Process Each Facility
# =====================

def process_facility(input_csv_path):
    filename = os.path.basename(input_csv_path)
    facility_name = clean_facility_name(filename)
    cocid = extract_cocid_from_filename(filename)
    print(f"Processing {facility_name}...")

    df = pd.read_csv(input_csv_path)
    if df.empty:
        print(f"⚠️ Skipped {facility_name}: empty file.")
        return

    # Remove duplicates
    possible_cols = ['ob_product_name', 'ib_product_name', 'Vendor', 'ib_vendor_name', 'ob_vendor_name', 'ib_datatype', 'ob_datatype']
    cols_for_dedup = [c for c in possible_cols if c in df.columns]
    if cols_for_dedup:
        df[cols_for_dedup] = df[cols_for_dedup].astype(str)
        df = df.drop_duplicates(subset=cols_for_dedup)

    # Apply special mapping for iPeople cases
    df['special'] = df.apply(get_special_ipeople_values, axis=1)
    df['Message Type'] = df['special'].apply(lambda x: x['MessageType'])
    df['Data Type'] = df['special'].apply(lambda x: x['DataType'])

    # Build final dataframe
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
        'Message Type': df['Message Type'],
        'Data Type': df['Data Type']
    }).dropna(subset=['Title 1', 'Title 2'])

    # Sort alphabetically
    filtered_df = filtered_df.sort_values(by=['Title 1', 'Title 2'])

    # Group by product and add bold header rows
    output_rows = []
    seen_products = set()
    for (title1, title2), group in filtered_df.groupby(['Title 1', 'Title 2'], sort=False):
        product_name = title1
        if product_name not in seen_products:
            product_row = {
                'Iteration Path': f"Integration-Tracker\\{wave_iteration}",
                'Area Path': f"Integration-Tracker\\{division}",
                'Work Item Type': 'Product',
                'Contract Ownership': '',
                'Product Owner': '',
                'Title 1': product_name,
                'Title 2': '',
                'State': "1. Contracting In Progress",
                'Facility': facility_name,
                'Facility Type': get_facility_type(facility_name),
                'COCID': cocid,
                'HCIS': default_hcis,
                'Division': division,
                'Product': product_name,
                'Connection Type': group['Connection Type'].iloc[0],
                'Data Flow Direction': '',
                'Expanse Interface Mnemonic': '',
                'Message Type': '',
                'Data Type': ''
            }
            output_rows.append(product_row)
            seen_products.add(product_name)
        for _, row in group.iterrows():
            r = row.copy()
            r['Title 1'] = ''
            output_rows.append(r)

    final_df = pd.DataFrame(output_rows)

    # Export to Excel
    output_excel_path = os.path.join(filtered_folder, f"{cocid}_{facility_name.replace(' ', '_')}.xlsx")
    final_df.to_excel(output_excel_path, index=False)

    # Style product rows (bold + centered)
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
    print(f"✅ {facility_name} exported successfully.\n")

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
