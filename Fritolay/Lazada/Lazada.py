import pandas as pd
import os
import shutil
import glob

# store_name = 'LAZADA PHILIPPINES (LAZADA GLICO)'

# Define parent directory
parent_dir = 'Fritolay'

# Define directories
raw_data_dir = os.path.join(parent_dir, 'Lazada', 'Inbound', 'RawData')
sku_dir = os.path.join(parent_dir, 'Lazada', 'Inbound', 'SKU')
consol_order_report_dir = os.path.join(parent_dir, 'Lazada', 'Inbound', 'ConsolOrderReport')
merged_dir = os.path.join(parent_dir, 'Lazada', 'Inbound', 'Merged')
archive_merged_dir = os.path.join(parent_dir, 'Tiktok', 'Archive', 'Merged')

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

# Function to clean SKU Reference No.
def clean_sku(sku):
    if sku in regular_special_skus:
        return sku
    if sku.startswith("STNG_") and "x" in sku:
        parts = sku.split("x")
        return f"ST_SBRYx{parts[1]}"
    if sku in specific_skus_to_clean:
        return sku.replace(specific_skus_to_clean[sku], "")
    for remove_str in skus_to_remove:
        sku = sku.replace(remove_str, "")
    return sku.split('x')[0] if 'x' in sku else sku

# Function to extract quantity from Seller SKU
def extract_quantity(seller_sku):
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
    if 'x' in seller_sku:
        try:
            return int(seller_sku.split('x')[1])
        except ValueError:
            return 1
    return 1

# Function to merge data from different directories
def merge_data(raw_data_dir, sku_dir, consol_order_report_dir, merged_dir):
    # Get list of .xlsx files in raw data directory
    raw_data_files = glob.glob(os.path.join(raw_data_dir, '*.xlsx'))
    
    # Merge files
    for raw_data_file in raw_data_files:
        raw_data = pd.read_excel(raw_data_file)

        # Clean sellerSku in raw_data
        raw_data['Cleaned SKU'] = raw_data['sellerSku'].apply(clean_sku)

        # Get list of .xlsx and .xls files in consol order report directory
        consol_order_report_files = glob.glob(os.path.join(consol_order_report_dir, '*.xlsx'))
        consol_order_report_files += glob.glob(os.path.join(consol_order_report_dir, '*.xls'))

        # Read all ConsolOrderReport files into a list of dataframes
        consol_order_report_dfs = []
        for consol_order_report_file in consol_order_report_files:
            df = pd.read_excel(consol_order_report_file)
            consol_order_report_dfs.append(df)

        # Concatenate all ConsolOrderReport dataframes into one
        consol_order_report = pd.concat(consol_order_report_dfs, ignore_index=True)

        # Filter consol_order_report for the desired Order Source
        consol_order_report_filtered = consol_order_report[
            consol_order_report['Order Source'] == 'Lazada Philippines (Lazada Frito-Lay)'
        ]

        # Convert 'orderNumber' and 'Order Number.' to string type
        raw_data['orderNumber'] = raw_data['orderNumber'].astype(str)
        consol_order_report_filtered['Order Number.'] = consol_order_report_filtered['Order Number.'].astype(str)

        # Convert 'Product Sku' to string type to match with 'Cleaned SKU'
        raw_data['Cleaned SKU'] = raw_data['Cleaned SKU'].astype(str)
        consol_order_report_filtered['Product Sku'] = consol_order_report_filtered['Product Sku'].astype(str)

        # Perform left join with RawData as base DataFrame on both orderNumber and Cleaned SKU/Product Sku
        # merged_data = pd.merge(raw_data, consol_order_report_filtered, how='left', 
        #                        left_on=['orderNumber', 'Cleaned SKU'], 
        #                        right_on=['Order Number.', 'Product Sku'])

        # Create a dictionary to map 'Order Number.' to 'Out of Warehouse' using the first occurrence
        order_warehouse_mapping = (
            consol_order_report_filtered[['Order Number.', 'Out of Warehouse']]
            .drop_duplicates('Order Number.')
            .set_index('Order Number.')
            .to_dict()['Out of Warehouse']
        )

        # Map 'Out of Warehouse' to raw_data using the dictionary
        raw_data['Out of Warehouse'] = raw_data['orderNumber'].map(order_warehouse_mapping)
        
        # Create new column "Qty" based on "sellerSku"
        raw_data['Qty'] = raw_data['sellerSku'].apply(extract_quantity)
        # merged_data['Qty'] = merged_data['Qty'] * merged_data['Quantity']

        # Remove 'x' and any digits after 'x' in sellerSku
        raw_data['sellerSku'] = raw_data['sellerSku'].apply(clean_sku)

        # Drop rows with duplicate "orderNumber"
        # merged_data = merged_data.drop_duplicates(subset=['orderNumber', 'Cleaned SKU', 'Qty'], keep='first')

        # Get list of .xlsx files in SKU directory
        sku_files = glob.glob(os.path.join(sku_dir, '*.xlsx'))
        for sku_file in sku_files:
            sku_data = pd.read_excel(sku_file)
            # Perform left join with RawData as base DataFrame
            raw_data['sellerSku'] = raw_data['sellerSku'].astype(str)
            sku_data['BRI MATCODE'] = sku_data['BRI MATCODE'].astype(str)
            raw_data = pd.merge(raw_data, sku_data, how='left', left_on='sellerSku', right_on='BRI MATCODE')
        
        # Ensure 'orderNumber' is of type str
        raw_data['orderNumber'] = raw_data['orderNumber'].astype(str)
        
        # Generate filename
        filename = os.path.basename(raw_data_file).replace(".xlsx", "_merged.xlsx")
        # Save merged data
        merged_path = os.path.join(merged_dir, filename)
        raw_data.to_excel(merged_path, index=False)
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
        merge_data['SC SALES'] = merge_data['unitPrice'] * merge_data['Qty']
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

        # Add a column for GROSS SALES filled with None
        # merge_data['GROSS SALES'] = None
        # merge_data['SC SALES'] = None
        # merge_data['COGS PRICE'] = None
        # merge_data['Voucher discounts'] = None
        # merge_data['Promo Discounts'] = None
        
        # Add a column for GROSS SALES filled with None
        # merge_data['GROSS SALES'] = None
        # merge_data['SC SALES'] = None
        # merge_data['COGS PRICE'] = None
        # merge_data['Voucher discounts'] = None
        # merge_data['Promo Discounts'] = None
        # merge_data['Other Income'] = None
        # merge_data['UDS'] = None
        # merge_data['PAID/UNPAID'] = None
        # merge_data['SALES'] = None
        # merge_data['PAYMENT'] = None
        # merge_data['Variance'] = None
        # merge_data['%'] = None

        # Rename columns and reorder
        merge_data = merge_data.rename(columns={
            'trackingCode': 'Trucking #',
            'orderNumber': 'ORDER ID',
            'sellerSku': 'Material No.',
            'Qty': 'Qty',
            'createTime': 'Order Creation Date',
            'unitPrice': 'SC Unit Price',
            # 'itemName': 'Material Description',
            'MATERIAL DESCRIPTION': 'Material Description',
            'sellerDiscountTotal': 'Voucher discounts',
            'orderNumber': 'ORDER ID',
            'status': 'DELIVERY STATUS',
            'Out of Warehouse': 'DISPATCH DATE',
            'wareHouse': 'wareHouse',
            'buyerFailedDeliveryReason': 'Cancelled Reason',
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
consolidation_dir = os.path.join(parent_dir, 'Lazada', 'Outbound', 'Consolidation')

# Call the function
generate_consolidation(merged_dir, consolidation_dir)

def generate_quickbook_upload(consolidation_dir, quickbooks_dir):
    # Find any xlsx files in the input directory
    input_files = glob.glob(os.path.join(consolidation_dir, '*.xlsx'))
    store_name = 'LAZADA PHILIPPINES (LAZADA FRITO-LAY)'

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
        })[['*InvoiceNo', '*Customer', '*InvoiceDate', 'DISPATCHED DATE + 30 DAYS', 'Terms', 'Location', 'Memo', 'Item(Product/Service)', 'ItemDescription', 'ItemQuantity', 'ItemRate', '*ItemTaxCode', 'ItemTaxAmount', 'Currency', 'Service Date']]
        

        # Convert '*InvoiceDate' to datetime if it's not already
        if not pd.api.types.is_datetime64_any_dtype(merge_data['*InvoiceDate']):
            merge_data['*InvoiceDate'] = pd.to_datetime(merge_data['*InvoiceDate'], errors='coerce')
        
        # Format the '*InvoiceDate' column to MM/DD/YYYY format
        merge_data['*InvoiceDate'] = merge_data['*InvoiceDate'].dt.strftime('%m/%d/%Y')
        
        # Format the '*InvoiceNo' column to include the date in MMDDYYYY format
        # merge_data['*InvoiceNo'] = merge_data['*InvoiceNo'] + merge_data['*InvoiceDate'].str.replace('/', '')

        # Ensure '*InvoiceNo' is of type str
        merge_data['*InvoiceNo'] = merge_data['*InvoiceNo'].astype(str)

        # Generate filename
        filename = os.path.basename(input_file).replace(".xlsx", "_quickbooks_upload.xlsx")
        
        # Save the modified data to the output directory
        output_path = os.path.join(quickbooks_dir, filename)
        merge_data.to_excel(output_path, index=False)
        print(f"Consolidation generated and saved to: {output_path}")

# Define directories
consolidation_dir = os.path.join(parent_dir, 'Lazada', 'Outbound', 'Consolidation')
quickbooks_dir = os.path.join(parent_dir, 'Lazada', 'Outbound', 'QuickBooks')

# Call the function
generate_quickbook_upload(consolidation_dir, quickbooks_dir)
