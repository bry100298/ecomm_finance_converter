import pandas as pd
from datetime import datetime
import os

# Define parent directory
parent_dir = 'Lazada'

# Define the consolidation fields mapping
consolidation_fields = {
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

def generate_consolidation_dataframe(consolidation_fields, raw_data):
    # Create DataFrame with consolidation fields
    consolidation_df = pd.DataFrame(columns=consolidation_fields.keys())

    # Iterate over consolidation fields and populate DataFrame with raw data
    for column, raw_column in consolidation_fields.items():
        if raw_column is not None and raw_column in raw_data.columns:
            if column == 'Material No.' or column == 'Qty':
                consolidation_df[column] = raw_data['sellerSku'].apply(lambda x: x.split('x')[0] if column == 'Material No.' else (x.split('x')[1] if 'x' in x else 1))
            else:
                consolidation_df[column] = raw_data[raw_column]
        else:
            consolidation_df[column] = None
    
    return consolidation_df

def process_raw_data(raw_data_file):
    # Read raw data from file
    raw_data = pd.read_excel(raw_data_file)
    return raw_data

# Define the directory containing raw data files
raw_data_dir = os.path.join(parent_dir, 'Inbound', 'RawData')

# Define the outbound directory
outbound_dir = os.path.join(parent_dir, 'Outbound')

# Process each raw data file
for root, dirs, files in os.walk(raw_data_dir):
    for file in files:
        if file.endswith('.xlsx'):
            raw_data_file = os.path.join(root, file)
            # Process raw data
            raw_data = process_raw_data(raw_data_file)
            # Generate consolidation DataFrame
            consolidation_df = generate_consolidation_dataframe(consolidation_fields, raw_data)
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            filename = f"consolidation_{timestamp}.xlsx"
            # Save consolidation DataFrame to Excel file in the outbound directory
            consolidation_df.to_excel(os.path.join(outbound_dir, filename), index=False)
            print(f"Processed {file} and saved consolidated data to {filename}")
