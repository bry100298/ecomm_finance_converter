import pandas as pd
import os
import shutil
import glob

# Define parent directory
parent_dir = 'Fritolay'

# Define directories
raw_data_dir = os.path.join(parent_dir, 'Tiktok', 'Inbound', 'RawData')
sku_dir = os.path.join(parent_dir, 'Tiktok', 'Inbound', 'SKU')
consol_order_report_dir = os.path.join(parent_dir, 'Tiktok', 'Inbound', 'ConsolOrderReport')
merged_dir = os.path.join(parent_dir, 'Tiktok', 'Inbound', 'Merged')
archive_merged_dir = os.path.join(parent_dir, 'Tiktok', 'Archive', 'Merged')
archive_raw_dir = os.path.join(parent_dir, 'Tiktok', 'Archive', 'RawData')

# Define special SKUs
regular_special_skus = {
    "PEP001X2+700033_1",
    "GAT001X6_1",
    "GATNS1X6_002",
    "700022x2+700024+700023+700785FG",
    "700015x2+700023+700021+700785",
    "MD002_CODMx2+120810_CODM",
    "120812_CODMx2+MD002_CODM",
    "120769x12 + 700785",
    "120767x12 + 700785",
    "120768x12 + 700785",
    "120812_CODMx2+MD001_CODM",
    "120802-05x2",
    "120804x10_BD",
    "120803x10_BD",
    "120805x10_BD",
    "120802x10_BD",
    "PCBOX01",
    "GATNS_002x3+GAT001x3",
    "MAXVOUCHERFLFG",
    "700016+120865+120861+700787x2",
    "XBDORITOSPACK",
    "700785FGLx2",
    "701977FGx2",
    "700787FGLx2",
    "LPTN_2X",
    "120835x2+33+34",
    "NBA-SQUAD120861x2+120891x2",
    "NBA-DUO120870x2+120862+120897FG",
    "NBA-FULL120865x2+120891x2+120862+700785FG+702055FG"
}

# Define SKUs to be cleaned by removing specific substrings
skus_to_remove = {
    "CS-KIT",
    "_ML",
    "LT_FG",
    "LT_FG1",
    "_NBA",
    "GP-",
    "ML_",
    "CBCS",
    "CSV2",
    "CSNEW",
}

# Define specific SKUs with descriptions to remove
specific_skus_to_clean = {
    "700029CS": "CS",
    "700795CS": "CS",
    "700789CS": "CS",
    "164200CS": "CS",
    "120774CS": "CS",
    "164383CS": "CS",
    "164379CS": "CS",
    "162659CS": "CS",
    "163754CS": "CS",
    "120690CS": "CS",
    "164386CS": "CS",
    "164384CS": "CS",
    "120798CS": "CS",
    "164196CS": "CS",
    "120180CS": "CS",
    "700785ICECS": "CS",
    "700022CS2": "CS2",
    "700067CS": "CS",
    "700009CS": "CS",
    "120820CS": "CS",
    "120807CS": "CS",
    "163861CS": "CS",
    "700012CS": "CS",
    "700014CS": "CS",
    "120828CS": "CS",
    "120818CS": "CS",
    "700788CS": "CS",
    "700065CS": "CS",
    "700023CS": "CS",
    "700008CS": "CS",
    "700025CS": "CS",
    "163862CS": "CS",
    "120861CS": "CS",
    "164655CS": "CS",
    "120804CS": "CS",
    "120857CS": "CS",
    "120870CS": "CS",
    "120865CS": "CS",
    "120840CS": "CS",
    "120839CS": "CS",
    "PEP320CS": "CS",
    "120426CS": "CS",
    "701976CS": "CS",
    "700782CS": "CS",
    "700784CS": "CS",
    "CS_120818": "CS_",
    "CS_700023": "CS_",
    "CS_701681": "CS_",
    "CS_120857": "CS_",
    "120890CS": "CS",
    "702094CS": "CS",
    "164480CS": "CS",
    "702095CS": "CS",
    "164793CS": "CS",
    "120894CS": "CS",
    "120892CS": "CS",
    "702067CS": "CS",
    "120861CBCSx2": "CBCSx2",
    "120892CSx2": "CSx2",
    "163863CS": "CS",
    "120044_1":"_1",
    "120809_CODM":"_CODM",
    "120811_CODM":"_CODM",
    "120812_CODM":"_CODM",
    "120828_CODM":"_CODM",
    "120810_CODM":"_CODM",
    "700788_CODM":"_CODM",
    "700785_CODM":"_CODM",
    "MD002_CODM":"_CODM",
    "MD001_CODM":"_CODM",
    "GP_120868":"GP_",
    "GP_120873":"GP_",
    "GP-120874":"GP_",
    "GP_120867":"GP_",
    "GP_120810":"GP_",
    "120864CB":"CB",
    "120870CB":"CB",
    "120862CB":"CB",
    "120861CB":"CB",
    "120865CB":"CB",
    "LPTN_RED_FG":"_FG"
}

specific_sku_clean_KTR = {
    "120800": "FG",
    "120825": "_2",
    "120804": "_2",
    "120865": "+FG_Snickers51G",
    "120861": "+FG_MMChocolate",
    "120870": "+FG_MMPeanut",
    "702162": "+FG_LPTNTMBLR",
}


# Function to extract quantity from Seller SKU
# def extract_quantity(seller_sku):
#     if seller_sku in regular_special_skus:
#         return 1
#     # if 'x' in seller_sku:
#     #     return int(seller_sku.split('x')[1])
#     if 'x' in seller_sku:
#         try:
#             return int(seller_sku.split('x')[1])
#         except ValueError:
#             return 1
#     else:
#         return 1

# # Function to clean SKU Reference No.
# def clean_sku(sku):
#     if sku in regular_special_skus:
#         return sku
#     if sku.startswith("STNG_") and "x" in sku:
#         parts = sku.split("x")
#         return f"ST_SBRYx{parts[1]}"
#     return sku.split('x')[0] if 'x' in sku else sku

# Function to clean SKU Reference No.
def clean_sku(sku):
    sku = str(sku)  # Convert sku to string
    if sku in regular_special_skus:
        return sku
    if sku.startswith("STNG_") and "x" in sku:
        parts = sku.split("x")
        return f"ST_SBRYx{parts[1]}"
    if sku in specific_skus_to_clean:
        return sku.replace(specific_skus_to_clean[sku], "")
    for remove_str in skus_to_remove:
        sku = sku.replace(remove_str, "")
    for sku_key in specific_sku_clean_KTR:
        if sku.startswith(sku_key):
            sku = sku.split('x')[0]  # Keep only the part before 'x'
            break
    return sku.split('x')[0] if 'x' in sku else sku


# Function to extract quantity from Seller SKU
def extract_quantity(seller_sku):
    seller_sku = str(seller_sku)  # Convert sku to string
    if seller_sku in regular_special_skus:
        return 1
    if 'ST_SBRYx' in seller_sku:
        try:
            return int(seller_sku.split('ST_SBRYx')[1])
        except ValueError:
            return 1
    if seller_sku in specific_skus_to_clean:
        return 1
    for remove_str in skus_to_remove:
        seller_sku = seller_sku.replace(remove_str, "")
    for sku_key in specific_sku_clean_KTR:
        if seller_sku.startswith(sku_key):
            # Extract quantity for specific_sku_clean_KTR
            try:
                return int(seller_sku.split(specific_sku_clean_KTR[sku_key])[0].split('x')[1])
            except (IndexError, ValueError):
                return 1
    # if 'x' in seller_sku:
    #     try:
    #         quantity_part = seller_sku.split('x')[1]
    #         # Keep only the numeric part of quantity_part
    #         quantity_part = ''.join(char for char in quantity_part if char.isdigit())
    #         return int(quantity_part)
    #     except ValueError:
    #         return 1
    return 1
    # if 'x' in seller_sku:
    #     try:
    #         return int(seller_sku.split('x')[1])
    #     except ValueError:
    #         return 1
    # return 1

# Function to merge data from different directories
def merge_data(raw_data_dir, sku_dir, consol_order_report_dir, merged_dir):
    # Get list of .xlsx files in raw data directory
    raw_data_files = glob.glob(os.path.join(raw_data_dir, '*.xlsx'))
    
    # Merge files
    for raw_data_file in raw_data_files:
        # Read raw data while skipping the second row (index 1)
        raw_data = pd.read_excel(raw_data_file, skiprows=lambda x: x in [1])

        # Get list of .xlsx and .xls files in consol order report directory
        # consol_order_report_files = glob.glob(os.path.join(consol_order_report_dir, '*.xlsx'))
        # consol_order_report_files += glob.glob(os.path.join(consol_order_report_dir, '*.xls'))

        # Get list of .xlsx and .xls files in consol order report directory
        consol_order_report_files = glob.glob(os.path.join(consol_order_report_dir, '*.xlsx')) + glob.glob(os.path.join(consol_order_report_dir, '*.xls'))

        # Read all ConsolOrderReport files into a list of dataframes
        consol_order_report_dfs = []
        for consol_order_report_file in consol_order_report_files:
            df = pd.read_excel(consol_order_report_file)
            consol_order_report_dfs.append(df)

        # Concatenate all ConsolOrderReport dataframes into one
        consol_order_report = pd.concat(consol_order_report_dfs, ignore_index=True)

        # Filter consol_order_report for the desired Order Source
        consol_order_report_filtered = consol_order_report[
            consol_order_report['Order Source'] == 'Tiktok Philippines (Tiktok Philippines - PEPSICO)'
        ]

        # Convert 'Order ID' and 'Order Number.' to string type
        raw_data['Order ID'] = raw_data['Order ID'].astype(str)
        consol_order_report_filtered['Order Number.'] = consol_order_report_filtered['Order Number.'].astype(str)

        consol_order_report_filtered['Product Sku'] = consol_order_report_filtered['Product Sku'].astype(str)

        order_warehouse_mapping = (
            consol_order_report_filtered[['Order Number.', 'Out of Warehouse']]
            .drop_duplicates('Order Number.')
            .set_index('Order Number.')
            .to_dict()['Out of Warehouse']
        )

        # Map 'Out of Warehouse' to raw_data using the dictionary
        raw_data['Out of Warehouse'] = raw_data['Order ID'].map(order_warehouse_mapping)

        
        # Create new column "Qty" based on "Seller SKU"
        raw_data['Qty'] = raw_data['Seller SKU'].apply(extract_quantity)
        # Ensure Quantity is of integer type
        # raw_data['Quantity'] = raw_data['Quantity'].astype(int)
        raw_data['Qty'] = raw_data['Qty'] * raw_data['Quantity']

        # Remove 'x' and any digits after 'x' in Seller SKU
        raw_data['Seller SKU'] = raw_data['Seller SKU'].apply(clean_sku)

        # Get list of .xlsx files in SKU directory
        sku_files = glob.glob(os.path.join(sku_dir, '*.xlsx'))
        for sku_file in sku_files:
            sku_data = pd.read_excel(sku_file)
            # Perform left join with RawData as base DataFrame
            raw_data['Seller SKU'] = raw_data['Seller SKU'].astype(str)
            sku_data['BRI MATCODE'] = sku_data['BRI MATCODE'].astype(str)
            raw_data = pd.merge(raw_data, sku_data, how='left', left_on='Seller SKU', right_on='BRI MATCODE')

        # Ensure 'orderNumber' is of type str
        raw_data['Order ID'] = raw_data['Order ID'].astype(str)
        
        # Generate filename
        filename = os.path.basename(raw_data_file).replace(".xlsx", "_merged.xlsx")
        # Save merged data
        merged_path = os.path.join(merged_dir, filename)
        raw_data.to_excel(merged_path, index=False)
        shutil.move(raw_data_file, os.path.join(archive_raw_dir, os.path.basename(raw_data_file)))
        print(f"Merged data from {os.path.basename(raw_data_file)}, ConsolOrderReport, and SKU and saved to {merged_path}")

# Merge data from different directories
merge_data(raw_data_dir, sku_dir, consol_order_report_dir, merged_dir)

def generate_consolidation(input_dir, output_dir):
    # Find any xlsx files in the input directory
    input_files = glob.glob(os.path.join(input_dir, '*.xlsx'))
    
    for input_file in input_files:
        # Read the Excel file
        merge_data = pd.read_excel(input_file)
        
        # Calculate GROSS SALES based on BRI SELLING PRICE (SRP) and Qty
        merge_data['GROSS SALES'] = merge_data['BRI SELLING PRICE (SRP)'] * merge_data['Qty']
        # Clean and convert 'SKU Seller Discount' to numeric
        merge_data['SKU Seller Discount'] = merge_data['SKU Seller Discount'].str.replace('PHP', '').str.replace(',', '').astype(float)
        merge_data['SKU Unit Original Price'] = merge_data['SKU Unit Original Price'].str.replace('PHP', '').str.replace(',', '').astype(float)
        merge_data['SC Unit Price'] = (
            merge_data['SKU Unit Original Price'] - 
            (merge_data['SKU Seller Discount'] / merge_data['Qty'])
        )
        merge_data['SC SALES'] = merge_data['SC Unit Price'] * merge_data['Qty']
        merge_data['COGS PRICE'] = merge_data['COGS'] * merge_data['Qty']

        # Calculate Promo Discounts based on conditions
        for index, row in merge_data.iterrows():
            if row['GROSS SALES'] >= row['SC SALES']:
                merge_data.at[index, 'Promo Discounts'] = row['GROSS SALES'] - row['SC SALES']
            else:
                merge_data.at[index, 'Promo Discounts'] = 0

        # Calculate Other Income based on conditions
        for index, row in merge_data.iterrows():
            if row['GROSS SALES'] <= row['SC SALES']:
                merge_data.at[index, 'Other Income'] = row['SC SALES'] - row['GROSS SALES']
            else:
                merge_data.at[index, 'Other Income'] = 0

        # Convert 'Out of Warehouse' to the desired format mm/dd/yyyy HH:MM:SS
        merge_data['Out of Warehouse'] = pd.to_datetime(merge_data['Out of Warehouse'], format='%d-%m-%Y %H:%M:%S').dt.strftime('%m/%d/%Y %H:%M:%S')
        
        merge_data['Voucher discounts'] = None
        # Calculate Voucher discounts
        # merge_data['Voucher discounts'] = merge_data.groupby('Order ID')['Seller Voucher(PHP)'].transform(lambda x: x / x.count())

        # Rename columns and reorder
        merge_data = merge_data.rename(columns={
            'Tracking ID': 'Trucking #',
            'Order ID': 'ORDER ID',
            'Seller SKU': 'Material No.',
            'Qty': 'Qty',
            'Created Time': 'Order Creation Date',
            # 'Deal Price': 'SC Unit Price',
            # 'Product Name': 'Material Description',
            'MATERIAL DESCRIPTION': 'Material Description',
            # 'sellerDiscountTotal': 'Voucher discounts',
            'Order ID': 'ORDER ID',
            'Order Status': 'DELIVERY STATUS',
            'Out of Warehouse': 'DISPATCH DATE',
            'Warehouse Name': 'wareHouse',
            'Cancel Reason': 'Cancelled Reason',
        })[['Trucking #', 'ORDER ID', 'Material No.', 'Qty', 'Order Creation Date', 'SC Unit Price', 'Material Description', 'GROSS SALES', 'SC SALES', 'COGS PRICE', 'Voucher discounts', 'Promo Discounts', 'Other Income', 'ORDER ID', 'DELIVERY STATUS', 'DISPATCH DATE', 'wareHouse', 'Cancelled Reason']]
        
        # Fill NaN values with 0 in specific columns
        columns_to_fill = ['GROSS SALES', 'SC SALES', 'COGS PRICE', 'Voucher discounts']
        merge_data[columns_to_fill] = merge_data[columns_to_fill].fillna(0)

        # Ensure 'ORDER ID' is of type str
        merge_data['ORDER ID'] = merge_data['ORDER ID'].astype(str)

        # Generate filename
        filename = os.path.basename(input_file).replace(".xlsx", "_consolidated.xlsx")
        # Generate filename
        # filename = os.path.basename(input_file).replace(".xlsx", "_consolidated.xlsx")
        
        # Save the modified data to the output directory
        output_path = os.path.join(output_dir, filename)
        merge_data.to_excel(output_path, index=False)
        print(f"Consolidation generated and saved to: {output_path}")
        # Move Merged Excel file to Archive Merged Folder
        shutil.move(input_file, os.path.join(archive_merged_dir, os.path.basename(input_file)))


consolidation_dir = os.path.join(parent_dir, 'Tiktok', 'Outbound', 'Consolidation')

# Call the function
generate_consolidation(merged_dir, consolidation_dir)

def generate_quickbook_upload(consolidation_dir, quickbooks_dir):
    # Find any xlsx files in the input directory
    input_files = glob.glob(os.path.join(consolidation_dir, '*.xlsx'))
    store_name = 'TIKTOK PHILIPPINES (PEPSICO)'

    for input_file in input_files:
        # Read the Excel file
        merge_data = pd.read_excel(input_file)
        


        merge_data['*Customer'] = store_name

        # Convert 'DISPATCH DATE' to datetime if it's not already
        if not pd.api.types.is_datetime64_any_dtype(merge_data['DISPATCH DATE']):
            merge_data['DISPATCH DATE'] = pd.to_datetime(merge_data['DISPATCH DATE'], errors='coerce')

        # Set DISPATCHED DATE + 30 DAYS as None
        merge_data['DISPATCHED DATE + 30 DAYS'] = merge_data['DISPATCH DATE'] + pd.Timedelta(days=30)
        merge_data['DISPATCHED DATE + 30 DAYS'] = merge_data['DISPATCHED DATE + 30 DAYS'].dt.strftime('%m/%d/%Y')

        merge_data['Terms'] = None
        merge_data['Location'] = None

        merge_data['Memo'] = merge_data['DISPATCH DATE'].dt.strftime('%m/%Y')

        merge_data['ItemDescription'] = merge_data['Material Description']

        merge_data['ItemRate'] = None
        merge_data['*ItemTaxCode'] = "12% S"
        merge_data['ItemTaxAmount'] = merge_data['GROSS SALES'] / 1.12 * 0.12
        merge_data['Currency'] = None
        merge_data['Service Date'] = None

        
        # Rename columns and reorder
        merge_data = merge_data.rename(columns={
            'ORDER ID': '*InvoiceNo',
            'DISPATCH DATE': '*InvoiceDate',
            'Material Description': 'Item(Product/Service)',
            'Qty': 'ItemQuantity',
            'GROSS SALES': '*ItemAmount',
        })[['*InvoiceNo', '*Customer', '*InvoiceDate', 'DISPATCHED DATE + 30 DAYS', 'Terms', 'Location', 'Memo', 'Item(Product/Service)', 'ItemDescription', 'ItemQuantity', '*ItemAmount', 'ItemRate', '*ItemTaxCode', 'ItemTaxAmount', 'Currency', 'Service Date']]
        

        # Convert '*InvoiceDate' to datetime if it's not already
        if not pd.api.types.is_datetime64_any_dtype(merge_data['*InvoiceDate']):
            merge_data['*InvoiceDate'] = pd.to_datetime(merge_data['*InvoiceDate'], errors='coerce')

        #ORIGINAL 
        # # Format the '*InvoiceDate' column to MM/DD/YYYY format
        # merge_data['*InvoiceDate'] = merge_data['*InvoiceDate'].dt.strftime('%m/%d/%Y')
        
        # # Format the '*InvoiceNo' column to include the date in MMDDYYYY format
        # # merge_data['*InvoiceNo'] = merge_data['*InvoiceNo'] + merge_data['*InvoiceDate'].str.replace('/', '')

        # # Ensure '*InvoiceNo' is of type str
        # merge_data['*InvoiceNo'] = merge_data['*InvoiceNo'].astype(str)
        
        # # Generate filename
        # filename = os.path.basename(input_file).replace(".xlsx", "_quickbooks_upload.xlsx")
        
        # # Save the modified data to the output directory
        # output_path = os.path.join(quickbooks_dir, filename)
        # merge_data.to_excel(output_path, index=False)
        # print(f"Consolidation generated and saved to: {output_path}")

        # Separate rows with and without valid '*InvoiceDate'
        valid_invoice_date = merge_data[merge_data['*InvoiceDate'].notnull()]
        pending_invoice_date = merge_data[merge_data['*InvoiceDate'].isnull()]

        # Process valid invoice date entries
        if not valid_invoice_date.empty:
            # Format the '*InvoiceDate' column to MM/DD/YYYY format
            valid_invoice_date['*InvoiceDate'] = valid_invoice_date['*InvoiceDate'].dt.strftime('%m/%d/%Y')

            # Generate filename
            valid_filename = os.path.basename(input_file).replace(".xlsx", "_quickbooks_upload.xlsx")

            # Save the modified data to the output directory
            valid_output_path = os.path.join(quickbooks_dir, valid_filename)
            valid_invoice_date.to_excel(valid_output_path, index=False)
            print(f"Consolidation generated and saved to: {valid_output_path}")

        # Process pending invoice date entries
        if not pending_invoice_date.empty:
            pending_filename = os.path.basename(input_file).replace(".xlsx", "_pending_quickbooks_upload.xlsx")

            # Create the pending directory if it doesn't exist
            pending_dir = os.path.join(quickbooks_dir, 'pending')
            os.makedirs(pending_dir, exist_ok=True)

            # Save the pending data to the pending directory
            pending_output_path = os.path.join(pending_dir, pending_filename)
            pending_invoice_date.to_excel(pending_output_path, index=False)
            print(f"Pending consolidation generated and saved to: {pending_output_path}")

# Define directories
consolidation_dir = os.path.join(parent_dir, 'Tiktok', 'Outbound', 'Consolidation')
quickbooks_dir = os.path.join(parent_dir, 'Tiktok', 'Outbound', 'QuickBooks')

# Call the function
generate_quickbook_upload(consolidation_dir, quickbooks_dir)
