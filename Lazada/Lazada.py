import pandas as pd
import os
import glob

# Define directories
# raw_data_dir = r'C:\Users\User\Documents\Project\ecomm_finance_converter\Lazada\Inbound\RawData'
# sku_dir = r'C:\Users\User\Documents\Project\ecomm_finance_converter\Lazada\Inbound\SKU'
# consol_order_report_dir = r'C:\Users\User\Documents\Project\ecomm_finance_converter\Lazada\Inbound\ConsolOrderReport'
# merged_dir = r'C:\Users\User\Documents\Project\ecomm_finance_converter\Lazada\Inbound\Merged'


# Define parent directory
parent_dir = 'Lazada'

# Define directories
raw_data_dir = os.path.join(parent_dir, 'Inbound', 'RawData')
sku_dir = os.path.join(parent_dir, 'SKU')
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
        merged_data['sellerSku'] = merged_data['sellerSku'].str.replace('x', '')

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
        merged_data = pd.read_excel(input_file)
        
        # Add a column for GROSS SALES filled with None
        merged_data['GROSS SALES'] = None
        merged_data['SC SALES'] = None
        merged_data['COGS PRICE'] = None
        # merged_data['Voucher discounts'] = None
        merged_data['Promo Discounts'] = None
        merged_data['Other Income'] = None
        merged_data['UDS'] = None
        merged_data['PAID/UNPAID'] = None
        merged_data['SALES'] = None
        merged_data['PAYMENT'] = None
        merged_data['Variance'] = None
        merged_data['%'] = None

        # Rename columns and reorder
        merged_data = merged_data.rename(columns={
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
            'sellerNote': 'Remarks',
            'orderItemId': 'ORDER ID'
        })[['Trucking #', 'ORDER ID', 'Material No.', 'Qty', 'Order Creation Date', 'SC Unit Price', 'Material Description', 'GROSS SALES', 'SC SALES', 'COGS PRICE', 'Voucher discounts', 'Promo Discounts', 'Other Income', 'ORDER ID', 'DELIVERY STATUS', 'DISPATCH DATE', 'wareHouse', 'Cancelled Reason', 'UDS', 'PAID/UNPAID', 'Remarks', 'ORDER ID', 'SALES', 'PAYMENT', 'Variance', 'PAID/UNPAID', '%']]
        
        # Generate filename
        filename = os.path.basename(input_file).replace(".xlsx", "_consolidated.xlsx")
        
        # Save the modified data to the output directory
        output_path = os.path.join(output_dir, filename)
        merged_data.to_excel(output_path, index=False)
        print(f"Consolidation generated and saved to: {output_path}")

consolidation_dir = os.path.join(parent_dir, 'Outbound', 'Consolidation')

# Call the function
generate_consolidation(merged_dir, consolidation_dir)