# =====================
# 📦 Import required libraries
# =====================
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# =====================
# 📁 Set up fixed file paths
# =====================
unfiltered_folder = r'C:\Users\joc4126\Desktop\CSV_File\Unfiltered'
filtered_folder = r'C:\Users\joc4126\Desktop\CSV_File\Filtered'
crosswalk_path = r'C:\Users\joc4126\Desktop\CSV_File\CSV_File\Files_To_Import\Data_Type_Crosswalk_Final.xlsx'
faciliy_path = r'C:\Users\joc4126\Desktop\CSV_File\CSV_File\Files_To_Import\Locations.xlsx'
connection_path = r'C:\Users\joc4126\Desktop\CSV_File\CSV_File\Files_To_Import\Connection_Type_Crosswalk.xlsx'

# Make sure the output folder exists
os.makedirs(filtered_folder, exist_ok=True)

# =====================
# 📖 Load and inspect the crosswalk Excel file
# =====================
crosswalk_df = pd.read_excel(crosswalk_path, usecols="A:F")
facility_df = pd.read_excel(faciliy_path, usecols="B:C")
connection_df = pd.read_excel(connection_path, usecols="A:C")

# Clean column names to remove leading/trailing spaces
crosswalk_df.columns = crosswalk_df.columns.str.strip()
facility_df.columns = facility_df.columns.str.strip()
connection_df.columns = connection_df.columns.str.strip()

# 🔍 Print column names to verify exact spelling
print("🔍 Crosswalk columns:", crosswalk_df.columns.tolist())
print("🔍 Facility columns:", facility_df.columns.tolist())
print("🔍 Connection columns:", connection_df.columns.tolist())

# Optional: print column names for debugging
# print("Crosswalk columns:", crosswalk_df.columns.tolist())

# Create mapping dictionaries
message_type_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 4])) # Message Type
data_type_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 5])) # DataType
exp_int_mnem_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 2])) # Expanse Interface Mnemonic
#  direction_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 3]))  # Uncomment if needed later
scooter_expanse_data_type_map = dict(zip(crosswalk_df.iloc[:, 0], crosswalk_df.iloc[:, 1])) # Scooter Data Type (Expanse)

facility_name_series = facility_df.iloc[:, 0] # Site Names
facility_cocid_map = dict(zip(facility_df.iloc[:, 0], facility_df.iloc[:, 1])) # COCID

connection_type_map = dict(zip(connection_df.iloc[:, 0], connection_df.iloc[:, 1])) # Connection Type
contract_map = dict(zip(connection_df.iloc[: 0], connection_df.iloc[:, 2])) # Contract Ownership



# =====================
# 🛠 Define the main processing function
# =====================

# Check facility name to ensure that it is input correctly
def check_facility_name(series, facility_name):
        
    if facility_name in series:
        return str(series.iloc[0])
    else:
        return facility_name


def process_facility(facility_name, input_csv_path):
    df = pd.read_csv(input_csv_path)
    print("\n🔍 Input CSV columns:", df.columns.tolist())  # Debug print
    # Defensive deduplication: only use columns that exist
    # Use column names for deduplication, only if they exist
    possible_cols = [
        'ob_product_name', 'ib_product_name', 'Vendor', 'ib_vendor_name', 'ob_vendor_name', 'ib_datatype', 'ob_datatype'
    ]
    cols_for_dedup = [col for col in possible_cols if col in df.columns]
    if cols_for_dedup:
        print(f"✅ Deduplication using columns: {cols_for_dedup}")
        # Convert deduplication columns to string for all rows
        try:
            for col in cols_for_dedup:
                df[col] = df[col].astype(str)
            df = df.drop_duplicates(subset=cols_for_dedup)
        except Exception as dedup_err:
            print(f"❌ Deduplication error: {dedup_err}. Skipping deduplication and continuing.")
    else:
        print("⚠️ No valid columns found for deduplication. Skipping deduplication.")

    facility_name = check_facility_name(facility_name_series, facility_name)

# Add Vendor string suffix removal    
# (insert code here)


    def get_data_flow(row):
        ib = row['ib_product_name']
        ob = row['ob_product_name']

        if ob == 'HCA 5.6' or 'MPF-Patient Folder/HPF Facility':
            return "Inbound"
        elif ib == 'HCA 5.6':
            return "Outbound"
        else:
            return ''


    def get_interface(row):
        data_type = data_type_map.get(row['ib_datatype'], '')
        direction = get_data_flow(row)
        if direction == 'Outbound':
            return f"{row['ob_vendor_name']} ({row['Vendor']}) - {data_type}"
        elif direction == 'Inbound':
            return f"{row['ib_vendor_name']} ({row['Vendor']}) - {data_type}"
        else:
            return ''

    def get_product(row):
        direction = get_data_flow(row)
        return row['ob_vendor_name'] if direction == 'Outbound' else row['ib_vendor_name']
    
    def get_connection_type(row):
#        combined = f"{row['ob_product_name']} {row['ib_product_name']}".lower()
#       if "kafka" in combined:
#            return "Kafka"
#        elif "waterpark" in combined:
#            return "Waterpark"
#        else:
#            return "Cloverleaf"

        # Defensive: handle missing 'Product' column
        product_val = row.get('Product', None)
        connection_check = connection_type_map.get(product_val, '')
        title2_val = row.get('Title 2', None)
        if title2_val == connection_check:
            return f"{connection_type_map.get(row.get('Connection_type', None))}"
        return connection_check

    
    def get_facilitytype(row):
        main_string = facility_name
        substring1 = 'Imaging'
        substring2 = 'Oncology'
        substring3 = 'Cancer'

        if substring1 in main_string:
            return "Imaging Center"
        elif substring2 or substring3 in main_string:
            return "Oncology Clinic"
        else:
            return "Acute"



    filtered_df = pd.DataFrame({        
        'Iteration Path': f"Integration-Tracker\{wave_iteration}",
        'Area Path': f"Integration-Tracker\{division}",
        'Work Item Type': 'Interface',
        'Contract Ownership': '',
        'Product Owner': '',
        'Title 1': df.apply(get_product, axis=1),
        'Title 2': df.apply(get_interface, axis=1),
        'State': "Contracting in Progress",
        'Facility': f"{facility_name}",
        'Facility Type': df.apply (get_facilitytype),
        'COCID': '',
        'HCIS': '',
        'Division': f"{division}",
        'Product': df.apply(get_product, axis=1),
        'Connection Type': df.apply(get_connection_type, axis=1),
        'Data Flow Direction': df.apply(get_data_flow, axis=1),
        'Expanse Interface Mnemonic': '',
        'Message Type': df['ib_datatype'].map(message_type_map),
        'Data Type': df['ib_datatype'].map(data_type_map)
    }).dropna(subset=['Title 1', 'Title 2'])

    filtered_df = filtered_df.sort_values(by=['Title 1', 'Title 2'], ascending=[True, True])

    output_rows = []
    for product, group in filtered_df.groupby(['Title 1','Title 2'], sort=False): # Changed groupby() logic to see if it will split by product AND vendor instead of just product.
        vendor_row = {
            'Iteration Path': f"Integration-Tracker\{wave_iteration}",
            'Area Path': f"Integration-Tracker\{division}",
            'Work Item Type': 'Product',
            'Contract Ownership': '',
            'Product Owner': '',
            'Title 1': f"{product}",
            'Title 2': '',
            'State': f"1. Contracting In Progress",
            'Facility': group['Facility'].iloc[0],
            'Facility Type': group['Facility Type'].iloc[0],
            'COCID': '',
            'HCIS': '',
            'Division': group['Division'].iloc[0],
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
        if ws.cell(row, 2).value == '':
            for col in range(1, ws.max_column + 1):
                ws.cell(row, col).font = bold_font
                ws.cell(row, col).alignment = center_align

    wb.save(output_excel_path)
    print(f"✅ Processed: {facility_name}")

# =====================
# 🔁 Batch loop to process all CSVs in Unfiltered
# =====================
wave_iteration = "Wave 2D"  # Set default or read from config if needed
division = "Gulf Coast Division"  # Set default or read from config if needed

# Process only the specified file in Unfiltered
filename = 'COCHRA_HCA_Florida_Highlands_Hospital.csv'
input_csv_path = os.path.join(unfiltered_folder, filename)
# Derive facility name from filename
base = filename.replace('.csv', '')
facility_name = base
for prefix in [
    'COCAT_', 'COCCL_', 'COCHRA_', 'COCKN_', 'COCMHB_', 'COCNW_', 'COCQR_', 'COCUH_', 'COCUHA_', 'COCWS_', 'COCXG_']:
    if facility_name.startswith(prefix):
        facility_name = facility_name[len(prefix):]
facility_name = facility_name.replace('-InterfaceMigrationReport', '')
facility_name = facility_name.replace('_', ' ').replace('-', ' ').strip()
print(f"Processing: {facility_name}")
try:
    process_facility(facility_name, input_csv_path)
except Exception as e:
    print(f"❌ Error processing {facility_name}: {e}")
