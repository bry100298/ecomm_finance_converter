import pandas as pd
from datetime import datetime
import os

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
