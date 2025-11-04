import pandas as pd
import os
import re
import argparse
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# =====================
# 📁 Set up fixed file paths
# =====================
unfiltered_folder = r'C:\Users\joc4126\Desktop\CSV_File\Unfiltered'
filtered_folder = r'C:\Users\joc4126\Desktop\CSV_File\Filtered'
crosswalk_path = r'C:\Users\joc4126\Desktop\CSV_File\CSV_File\Files_To_Import\Data_Type_Crosswalk_Final.xlsx'

os.makedirs(filtered_folder, exist_ok=True)

# =====================
# 📖 Load and validate the crosswalk Excel file
# =====================
# Read crosswalk and be tolerant to slightly different header names
crosswalk_df = pd.read_excel(crosswalk_path)
crosswalk_df.columns = crosswalk_df.columns.str.strip()
print("🔍 Crosswalk columns:", crosswalk_df.columns.tolist())

# Try several possible header names used in different versions of the file
magic_col = next((c for c in crosswalk_df.columns if 'Scooter Data Type (Magic)' in c), None)
expanse_col = next((c for c in crosswalk_df.columns if 'Scooter Data Type (Expanse)' in c), None)
mnemonic_col = next((c for c in crosswalk_df.columns if 'Expanse Interface Mnemonic' in c or 'DataType' in c or 'Data Type' in c), None)

if not magic_col or not expanse_col or not mnemonic_col:
    raise ValueError(f"❌ Crosswalk missing required columns. Found: {crosswalk_df.columns.tolist()}")

message_type_map = dict(zip(crosswalk_df[magic_col], crosswalk_df[expanse_col]))
data_type_map = dict(zip(crosswalk_df[magic_col], crosswalk_df[mnemonic_col]))

# =====================
# 🛠 Define the main processing function
# =====================
def process_file(input_csv_path):
    # Derive a readable facility name from the filename
    basename = os.path.basename(input_csv_path)
    # remove extension
    name = re.sub(r'(?i)\.csv$', '', basename)
    # remove common suffixes like "-InterfaceMigrationReport" (case-insensitive)
    name = re.sub(r'(?i)interface\s*migration\s*report', '', name)
    # remove leading site codes like "COCCL-" or "COCUH_" (uppercase letters/digits + separator)
    name = re.sub(r'^[A-Z0-9]+[_-]+', '', name)
    # replace underscores and dashes with spaces
    name = name.replace('_', ' ').replace('-', ' ').strip()
    # split on HCA token: if remaining contains 'HCA' and other words, prefer the HCA tail
    if 'HCA' in name:
        parts = name.split('HCA')
        tail = parts[-1].strip()
        if tail:
            name = f"HCA {tail}"
    else:
        # convert CamelCase words to space separated (e.g., HoustonClearLake -> Houston Clear Lake)
        name = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', name)
        # normalize multiple spaces
        name = re.sub(r'\s+', ' ', name).strip()
        # prefix with HCA for clarity
        name = f"HCA {name}"

    # Special-case: convert "HCA Houston Clear Lake" -> "HCA Houston Healthcare Clear Lake" if appropriate
    if re.search(r'HCA\s+Houston\s+Clear\s+Lake', name, flags=re.I):
        name = re.sub(r'HCA\s+Houston\s+Clear\s+Lake', 'HCA Houston Healthcare Clear Lake', name, flags=re.I)

    facility_name = name

    print(f"Processing file: {input_csv_path} -> Facility: {facility_name}")
    df = pd.read_csv(input_csv_path)

    # drop duplicates using a best-effort set of columns (if present)
    possible_idx = [2, 3, 4, 6, 9, 19]
    cols_for_dedup = [df.columns[i] for i in possible_idx if i < len(df.columns)]
    if cols_for_dedup:
        df = df.drop_duplicates(subset=cols_for_dedup)

    def get_connection_type(row):
        # tolerate missing columns
        ob = str(row.get('ob_product_name', '') or '')
        ib = str(row.get('ib_product_name', '') or '')
        combined = f"{ob} {ib}".lower()
        if "kafka" in combined:
            return "Kafka"
        elif "waterpark" in combined:
            return "Waterpark"
        else:
            return "Cloverleaf"

    def get_data_flow(row):
        # prefer explicit data flow column if present
        val = row.get('Data Flow Direction') or row.get('data_flow') or row.get('direction')
        if pd.notna(val) and str(val).strip():
            return str(val).strip().capitalize()
        # fallback to using ib_thread_type/ob_thread_type or vendor direction column
        # default to 'Inbound'
        return 'Inbound'

    def get_interface(row):
        ib_dt = row.get('ib_datatype') or row.get('ib_datatype'.lower()) or ''
        data_type = data_type_map.get(ib_dt, ib_dt)
        direction = get_data_flow(row)
        vendor_field = row.get('Vendor', '')
        if direction == 'Outbound':
            vendor = row.get('ob_vendor_name') or row.get('ob_vendor_mnem') or ''
            return f"{vendor} ({vendor_field}) - {data_type}"
        else:
            vendor = row.get('ib_vendor_name') or row.get('ib_vendor_mnem') or ''
            return f"{vendor} ({vendor_field}) - {data_type}"

    def get_product(row):
        direction = get_data_flow(row)
        return row.get('ob_vendor_name') if direction == 'Outbound' else row.get('ib_vendor_name')

    filtered_df = pd.DataFrame({
        'Product': df.apply(get_product, axis=1),
        'Interface': df.apply(get_interface, axis=1),
        'Facility': facility_name,
        'Connection Type': df.apply(get_connection_type, axis=1),
        'Data Flow Direction': df.apply(get_data_flow, axis=1),
        'Message Type': df.get('ib_datatype').map(message_type_map) if 'ib_datatype' in df.columns else df.get('ib_datatype'),
        'Data Type': df.get('ib_datatype').map(data_type_map) if 'ib_datatype' in df.columns else df.get('ib_datatype')
    }).dropna(subset=['Product', 'Interface'])

    filtered_df = filtered_df.sort_values(by=['Product', 'Interface'])

    output_rows = []
    for product, group in filtered_df.groupby('Product', sort=False):
        vendor_row = {
            'Product': f"{product}",
            'Interface': '',
            'Facility': group['Facility'].iloc[0],
            'Connection Type': group['Connection Type'].iloc[0],
            'Data Flow Direction': '',
            'Message Type': '',
            'Data Type': ''
        }
        output_rows.append(vendor_row)

        for _, row in group.iterrows():
            row_copy = row.copy()
            row_copy['Product'] = ''
            output_rows.append(row_copy)

    final_df = pd.DataFrame(output_rows)

    safe_name = facility_name.replace(' ', '_').replace('/', '_')
    # Also write a raw CSV that matches the rough column layout in your example screenshot.
    output_csv_path = os.path.join(filtered_folder, f"{safe_name}_raw.csv")

    export_rows = []
    current_product = None
    for _, r in final_df.iterrows():
        # r['Product'] is filled for vendor header rows; interface rows have Product == ''
        if str(r.get('Product') or '').strip():
            current_product = r['Product']
            work_type = 'Product'
            product_val = current_product
            title1 = ''
        else:
            work_type = 'Interface'
            product_val = current_product
            title1 = r.get('Interface', '')

        export_rows.append({
            'Work Item Type': work_type,
            'Contract Ownership': '',
            'Product Owner': '',
            'Title 1': title1,
            'Title 2': '',
            'Product': product_val,
            'Connection Type': r.get('Connection Type', ''),
            'Data Flow Direction': r.get('Data Flow Direction', ''),
            'Expanse Interface Mnemonic': r.get('Data Type', ''),
            'Message Type': r.get('Message Type', ''),
            'DataType': r.get('Data Type', '')
        })

    export_df = pd.DataFrame(export_rows)
    export_df.to_csv(output_csv_path, index=False)

    # Then write the formatted Excel as before
    output_excel_path = os.path.join(filtered_folder, f"{safe_name}.xlsx")
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
    print(f"✅ Interface data exported for: {facility_name}")

    # =====================
    # 🧱 Static product structure
    # =====================
    product_data = [
        {"Product ID": "Product: 10476", "Product": "Bhang - Blueberry"},
        {"Product ID": "Product: 10477", "Product": "Bhang - Caramel"},
        {"Product ID": "Product: 10478", "Product": "Bhang - Cherry"}
    ]

    product_columns = ["Product ID", "Product", "Title", "Facility", "Bin", "Slot", "Packet", "Type"]
    product_df = pd.DataFrame(product_data)
    for col in product_columns:
        if col not in product_df.columns:
            product_df[col] = ""

    # Use the cleaned facility_name directly for product structure (avoid duplicate HCA prefixes)
    product_df["Facility"] = facility_name
    product_df["Type"] = "Chemical"

    product_output_path = os.path.join(filtered_folder, f"Product_Structure_{safe_name}.xlsx")
    product_df.to_excel(product_output_path, index=False)
    print(f"📦 Product structure exported for: {facility_name}")

# =====================
# 🚀 Static facility list
# =====================
# Process all CSV files found in the Unfiltered folder
def main():
    parser = argparse.ArgumentParser(description='Process one Unfiltered CSV or all in folder')
    parser.add_argument('--file', '-f', help='Path to a single CSV file to process (optional)')
    args = parser.parse_args()

    if args.file:
        path = args.file
        if not os.path.exists(path):
            print(f"❌ Specified file does not exist: {path}")
            return
        try:
            process_file(path)
        except Exception as e:
            print(f"❌ Error processing {path}: {e}")
    else:
        csv_files = [os.path.join(unfiltered_folder, f) for f in os.listdir(unfiltered_folder) if f.lower().endswith('.csv')]
        if not csv_files:
            print(f"⚠️ No CSV files found in {unfiltered_folder}")
            return
        for path in csv_files:
            try:
                process_file(path)
            except Exception as e:
                print(f"❌ Error processing {path}: {e}")


if __name__ == '__main__':
    main()
