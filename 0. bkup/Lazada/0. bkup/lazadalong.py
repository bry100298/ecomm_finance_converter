import pandas as pd
import os
import glob

# Define directories
raw_data_dir = r'C:\Users\User\Documents\Project\ecomm_finance_converter\Lazada\Inbound\RawData'
sku_dir = r'C:\Users\User\Documents\Project\ecomm_finance_converter\Lazada\Inbound\SKU'
consol_order_report_dir = r'C:\Users\User\Documents\Project\ecomm_finance_converter\Lazada\Inbound\ConsolOrderReport'
merged_dir = r'C:\Users\User\Documents\Project\ecomm_finance_converter\Lazada\Inbound\Merged'

# Function to extract quantity from sellerSku
def extract_quantity(seller_sku):
    if 'x' in seller_sku:
        return int(seller_sku.split('x')[1])
    else:
        return 1

# Function to merge data from different directories
def merge_data(raw_data_dir, consol_order_report_dir, merged_dir):
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
        
        # Remove letter "x" from "sellerSku" and create a new column "sellerSku_clean"
        merged_data['sellerSku_clean'] = merged_data['sellerSku'].str.replace('x', '')
        merged_data['sellerSku_clean'] = merged_data['sellerSku_clean'].astype(str)
        
        # Drop rows with duplicate "orderItemId"
        merged_data = merged_data.drop_duplicates(subset='orderItemId', keep='first')

        # Generate filename
        filename = os.path.basename(raw_data_file).replace(".xlsx", "_merged.xlsx")
        # Save merged data
        merged_path = os.path.join(merged_dir, filename)
        merged_data.to_excel(merged_path, index=False)
        print(f"Merged data from {os.path.basename(raw_data_file)}, ConsolOrderReport, and saved to {merged_path}")
        
        # Return merged data
        return merged_data

# Function to merge SKU data
def merge_data_sku(sku_dir, merged_dir, merged_data):
    # Get list of .xlsx files in SKU directory
    sku_files = glob.glob(os.path.join(sku_dir, '*.xlsx'))
    for sku_file in sku_files:
        sku_data = pd.read_excel(sku_file)
        print("SKU Data:")
        print(sku_data.head())  # Print first few rows of SKU data for debugging
        
        # Convert 'sellerSku_clean' and 'BRI MATCODE' columns to string type
        merged_data['sellerSku_clean'] = merged_data['sellerSku_clean'].astype(str)
        sku_data['BRI MATCODE'] = sku_data['BRI MATCODE'].astype(str)
        # Perform left join with Merged data as base DataFrame using the 'sellerSku_clean' column
        merged_data = pd.merge(merged_data, sku_data, how='left', left_on='sellerSku_clean', right_on='BRI MATCODE')
        
        print("Merged Data:")
        print(merged_data.head())  # Print first few rows of merged data for debugging
    
    # Generate filename
    filename = "sku_merged.xlsx"
    # Save merged SKU data
    merged_path = os.path.join(merged_dir, "Outbound", filename)
    merged_data.to_excel(merged_path, index=False)
    print(f"Merged SKU data saved to {merged_path}")


# Merge data from different directories and get merged data
merged_data = merge_data(raw_data_dir, consol_order_report_dir, merged_dir)

# Merge SKU data using the obtained merged data
merge_data_sku(sku_dir, merged_dir, merged_data)
