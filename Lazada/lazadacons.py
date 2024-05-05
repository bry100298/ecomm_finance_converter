import pandas as pd
from datetime import datetime
import os
import glob

# Define parent directory
parent_dir = 'Lazada'

# Define the partial consolidation fields mapping
partial_consolidation_fields = {
    'Trucking #': 'trackingCode',
    'ORDER ID': 'orderItemId',
    'Material No.': 'sellerSku',
    'Qty': 'sellerSku',
    'Order Creation Date': 'createTime',
    'SC Unit Price': 'unitPrice',
    'Material Description': 'itemName',
    'GROSS SALES': None,
    'SC SALES': None,
    'COGS PRICE': None,
    'Voucher discounts': 'sellerDiscountTotal',
    'Promo Discounts': None,
    'Other Income': None,
    'ORDER ID': 'LazadaId',
    'DELIVERY STATUS': 'status',
    'DISPATCH DATE': None,
    'wareHouse': 'wareHouse',
    'Cancelled Reason': None,
    'UDS': None,
    'PAID/UNPAID': None,
    'Remarks': 'sellerNote',
    'ORDER ID': 'orderItemId',
    'SALES': None,
    'PAYMENT': None,
    'Variance': None,
    'PAID/UNPAID': None,
    '%': None,
    'Remarks': 'sellerNote'
}

def generate_partial_consolidation_dataframe(partial_consolidation_fields, raw_data):
    # Create DataFrame with partial consolidation fields
    partial_consolidation_df = pd.DataFrame(columns=partial_consolidation_fields.keys())

    # Iterate over partial consolidation fields and populate DataFrame with raw data
    for column, raw_column in partial_consolidation_fields.items():
        if raw_column is not None and raw_column in raw_data.columns:
            if column == 'Material No.' or column == 'Qty':
                partial_consolidation_df[['Material No.', 'Qty']] = raw_data['sellerSku'].str.split('x', n=1, expand=True)
                partial_consolidation_df['Qty'] = pd.to_numeric(partial_consolidation_df['Qty'], errors='coerce').fillna(1)
            else:
                partial_consolidation_df[column] = raw_data[raw_column]
        else:
            partial_consolidation_df[column] = None
    
    return partial_consolidation_df


def process_raw_data(raw_data_file):
    # Read raw data from file
    raw_data = pd.read_excel(raw_data_file)
    return raw_data

# Define the directory containing raw data files
raw_data_dir = os.path.join(parent_dir, 'Inbound', 'RawData')

# Define the outbound directory
outbound_dir = os.path.join(parent_dir, 'Outbound')

# Define the partial outbound directory
partial_outbound_dir = os.path.join(parent_dir, 'Inbound', 'ConsolOrderReport', 'Partial')



def merge_files(parent_dir):
    # Define directories
    partial_dir = os.path.join(parent_dir, 'Inbound', 'ConsolOrderReport', 'Partial')
    xls_dir = os.path.join(parent_dir, 'Inbound', 'ConsolOrderReport')
    merged_dir = os.path.join(parent_dir, 'Inbound', 'ConsolOrderReport', 'Merged')
    
    # Create merged directory if it doesn't exist
    if not os.path.exists(merged_dir):
        os.makedirs(merged_dir)
    
    # Get list of .xlsx files in the partial directory
    partial_files = glob.glob(os.path.join(partial_dir, '*.xlsx'))
    
    # Get list of .xls files in the xls directory
    xls_files = glob.glob(os.path.join(xls_dir, '*.xls'))
    
    # Print out the list of .xls files found
    print("List of .xls files found:")
    print(xls_files)
    
    # Iterate through each .xlsx file in the partial directory
    for partial_file in partial_files:
        # Load partial .xlsx file
        partial_df = pd.read_excel(partial_file)
        
        # Reset index of partial DataFrame
        partial_df.reset_index(drop=True, inplace=True)
        
        # Iterate through each .xls file in the xls directory
        for xls_file in xls_files:
            # Load .xls file
            xls_df = pd.read_excel(xls_file)
            
            # Check and convert data types if necessary for merging
            if partial_df['ORDER ID'].dtype != xls_df['Order Number.'].dtype:
                if partial_df['ORDER ID'].dtype == 'object':
                    xls_df['Order Number.'] = xls_df['Order Number.'].astype(str)
                else:
                    partial_df['ORDER ID'] = partial_df['ORDER ID'].astype(str)
            
            # Merge based on common column "ORDER ID" and "Order Number."
            merged_df = pd.merge(partial_df, xls_df, how='inner', left_on='ORDER ID', right_on='Order Number.')
            
            # Save merged DataFrame to .xlsx file in merged directory
            merged_filename = os.path.basename(partial_file).replace("LazadaPartialCons_", "LazadaPartialConsMerged_")
            merged_path = os.path.join(merged_dir, merged_filename)
            merged_df.to_excel(merged_path, index=False)
            print(f"Merged data from {os.path.basename(partial_file)} with {os.path.basename(xls_file)} and saved to {merged_path}")



# Define the new merged order report consolidation fields mapping
merged_order_report_consolidation_fields = {
    'Trucking #': 'trackingCode',
    'ORDER ID': 'orderItemId',
    'Material No.': 'sellerSku',
    'Qty': 'sellerSku',
    'Order Creation Date': 'createTime',
    'SC Unit Price': 'unitPrice',
    'Material Description': 'itemName',
    'GROSS SALES': None,
    'SC SALES': None,
    'COGS PRICE': None,
    'Voucher discounts': 'sellerDiscountTotal',
    'Promo Discounts': None,
    'Other Income': None,
    'ORDER ID': 'LazadaId',
    'DELIVERY STATUS': 'status',
    'DISPATCH DATE': 'Out of Warehouse',
    'wareHouse': 'wareHouse',
    'Cancelled Reason': 'Cancel Reason',
    'UDS': None,
    'PAID/UNPAID': None,
    'Remarks': 'sellerNote',
    'ORDER ID': 'orderItemId',
    'SALES': None,
    'PAYMENT': None,
    'Variance': None,
    'PAID/UNPAID': None,
    '%': None,
    'Remarks': 'sellerNote'
}

# Define the reduce outbound directory
reduce_outbound_dir = os.path.join(parent_dir, 'Inbound', 'ConsolOrderReport', 'Merged', 'Reduce')

# Create reduce outbound directory if it doesn't exist
if not os.path.exists(reduce_outbound_dir):
    os.makedirs(reduce_outbound_dir)

# Function to reduce the merged data
# Function to reduce the merged data
def reduce_merged_data(merged_dir):
    # Get list of .xlsx files in the merged directory
    merged_files = glob.glob(os.path.join(merged_dir, '*.xlsx'))
    
    # Iterate through each .xlsx file in the merged directory
    for merged_file in merged_files:
        # Read the merged data
        merged_data = pd.read_excel(merged_file)
        
        # Rename columns if necessary
        merged_data.rename(columns={
            'Out of Warehouse': 'DISPATCH DATE',
            'Cancel Reason': 'Cancelled Reason'
        }, inplace=True)
        
        # Select only the columns specified in merged_order_report_consolidation_fields
        merged_data_reduced = merged_data[list(merged_order_report_consolidation_fields.keys())]
        
        # Generate filename for reduced data
        reduce_filename = os.path.basename(merged_file).replace("LazadaPartialConsMerged_", "LazadaReduce_")
        reduce_path = os.path.join(reduce_outbound_dir, reduce_filename)
        
        # Save reduced data to .xlsx file in reduce outbound directory
        merged_data_reduced.to_excel(reduce_path, index=False)
        print(f"Reduced data from {os.path.basename(merged_file)} and saved to {reduce_path}")


# Process each raw data file
for root, dirs, files in os.walk(raw_data_dir):
    for file in files:
        if file.endswith('.xlsx'):
            raw_data_file = os.path.join(root, file)
            # Process raw data
            raw_data = process_raw_data(raw_data_file)
            # Generate partial consolidation DataFrame
            partial_consolidation_df = generate_partial_consolidation_dataframe(partial_consolidation_fields, raw_data)
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            filename = f"LazadaPartialCons_{timestamp}.xlsx"
            # Save partial consolidation DataFrame to Excel file in the partial outbound directory
            partial_outbound_path = os.path.join(partial_outbound_dir, filename)
            partial_consolidation_df.to_excel(partial_outbound_path, index=False)
            print(f"Processed {file} and saved partial consolidated data to {partial_outbound_path}")


# Call the merge_files function
merge_files(parent_dir)

# Define the merged directory
merged_dir = os.path.join(parent_dir, 'Inbound', 'ConsolOrderReport', 'Merged')

# Call the reduce_merged_data function
reduce_merged_data(merged_dir)