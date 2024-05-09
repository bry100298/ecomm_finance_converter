import pandas as pd
import os
import glob

# Define parent directory
parent_dir = 'Lazada'

# Define directories
raw_data_dir = os.path.join(parent_dir, 'Inbound', 'RawData')
sku_dir = os.path.join(parent_dir, 'Inbound', 'SKU')
consol_order_report_dir = os.path.join(parent_dir, 'Inbound', 'ConsolOrderReport')
merged_dir = os.path.join(parent_dir, 'Inbound', 'Merged')

# Function to extract quantity from sellerSku
def extract_quantity(seller_sku):
    if 'x' in seller_sku:
        return int(seller_sku.split('x')[1])
    else:
        return 1

# Function to merge data from different directories
def merge_data(raw_data_dir, sku_dir, consol_order_report_dir, merged_dir):
    # Get list of .xlsx files in raw data directory
    raw_data_files = glob.glob(os.path.join(raw_data_dir, '*.xlsx'))
    
    # Merge files
    for raw_data_file in raw_data_files:
        raw_data = pd.read_excel(raw_data_file)
        
        # Get list of .xls files in consol order report directory
        consol_order_report_files = glob.glob(os.path.join(consol_order_report_dir, '*.xls'))
        for consol_order_report_file in consol_order_report_files:
            consol_order_report = pd.read_excel(consol_order_report_file)
            # Perform left join with RawData as base DataFrame
            merged_data = pd.merge(raw_data, consol_order_report, how='left', left_on='orderItemId', right_on='Order Number.')
        
        # Create new column "Qty" based on "sellerSku"
        merged_data['Qty'] = merged_data['sellerSku'].apply(extract_quantity)
        
        # Remove letter "x" from "sellerSku"
        # merged_data['sellerSku'] = merged_data['sellerSku'].str.replace('x', '')
        # Remove 'x' and any digits after 'x' in sellerSku
        merged_data['sellerSku'] = merged_data['sellerSku'].apply(lambda x: x.split('x')[0] if 'x' in x else x)


        # Drop rows with duplicate "orderItemId"
        merged_data = merged_data.drop_duplicates(subset='orderItemId', keep='first')

        # Get list of .xlsx files in SKU directory
        sku_files = glob.glob(os.path.join(sku_dir, '*.xlsx'))
        for sku_file in sku_files:
            sku_data = pd.read_excel(sku_file)
            # Perform left join with RawData as base DataFrame
            merged_data['sellerSku'] = merged_data['sellerSku'].astype(str)
            sku_data['BRI MATCODE'] = sku_data['BRI MATCODE'].astype(str)
            merged_data = pd.merge(merged_data, sku_data, how='left', left_on='sellerSku', right_on='BRI MATCODE')
        
        # Generate filename
        filename = os.path.basename(raw_data_file).replace(".xlsx", "_merged.xlsx")
        # Save merged data
        merged_path = os.path.join(merged_dir, filename)
        merged_data.to_excel(merged_path, index=False)
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
            'orderItemId': 'ORDER ID',
            'sellerSku': 'Material No.',
            'Qty': 'Qty',
            'createTime': 'Order Creation Date',
            'unitPrice': 'SC Unit Price',
            'itemName': 'Material Description',
            'sellerDiscountTotal': 'Voucher discounts',
            'orderItemId': 'ORDER ID',
            'status': 'DELIVERY STATUS',
            'Out of Warehouse': 'DISPATCH DATE',
            'wareHouse': 'wareHouse',
            'buyerFailedDeliveryReason': 'Cancelled Reason',
        })[['Trucking #', 'ORDER ID', 'Material No.', 'Qty', 'Order Creation Date', 'SC Unit Price', 'Material Description', 'GROSS SALES', 'SC SALES', 'COGS PRICE', 'Voucher discounts', 'Promo Discounts', 'Other Income', 'ORDER ID', 'DELIVERY STATUS', 'DISPATCH DATE', 'wareHouse', 'Cancelled Reason']]
        

        # Fill NaN values with 0 in specific columns
        columns_to_fill = ['GROSS SALES', 'SC SALES', 'COGS PRICE', 'Voucher discounts']
        merge_data[columns_to_fill] = merge_data[columns_to_fill].fillna(0)

        # Generate filename
        filename = os.path.basename(input_file).replace(".xlsx", "_consolidated.xlsx")
        # Generate filename
        filename = os.path.basename(input_file).replace(".xlsx", "_consolidated.xlsx")
        
        # Save the modified data to the output directory
        output_path = os.path.join(output_dir, filename)
        merge_data.to_excel(output_path, index=False)
        print(f"Consolidation generated and saved to: {output_path}")

consolidation_dir = os.path.join(parent_dir, 'Outbound', 'Consolidation')

# Call the function
generate_consolidation(merged_dir, consolidation_dir)