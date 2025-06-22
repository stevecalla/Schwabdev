from dotenv import load_dotenv  # Import dotenv to load .env variables
import os  # Import os to access environment variables
import pandas as pd
import json

from schwabdev.client import Client  # Import the Client class from schwabdev/client.py
import schwabdev  # Import the schwabdev package to load the client module

# Load environment variables from .env file
load_dotenv()

# Access the environment variables for client_id and client_secret
client_id = os.getenv('APP_KEY')  # Get the client_id from .env
client_secret = os.getenv('APP_SECRET')  # Get the client_secret from .env

# Create a client using the credentials loaded from the .env file
client = Client(client_id, client_secret)  # Create an instance of the Client class

# Fetch account details with positions
response = client.account_details_all("positions").json()  # Fetch account details

# Pretty print the JSON response for debugging
# print("Response JSON (Pretty Printed):")
# print(json.dumps(response, indent=4))  # Pretty print the JSON response

# Flatten the data to a simple structure
flattened_data = []
for account in response:
    account_info = account['securitiesAccount']
    
    # For each position in the account's 'positions' list, we flatten the data
    for position in account_info.get('positions', []):
        # Safely access 'description' field with a fallback to 'None' if not found
        position_description = position['instrument'].get('description', 'N/A')  # Default to 'N/A' if description is missing
        
        flattened_data.append({
            'accountNumber': str(account_info['accountNumber']),  # Ensure account number is string with no commas or decimals
            'positionSymbol': position['instrument'].get('symbol', 'N/A'),  # Position's symbol (e.g., 'BAC', 'TSLA')
            'positionDescription': position_description,  # Position description
            'positionType': position['instrument'].get('type', 'N/A'),  # Position type (e.g., 'EQUITY', 'FIXED_INCOME')
            'shortQuantity': position.get('shortQuantity', 0),
            'longQuantity': position.get('longQuantity', 0),
            'averagePrice': position.get('averagePrice', 0),
            'currentDayProfitLoss': position.get('currentDayProfitLoss', 0),
            'currentDayProfitLossPercentage': position.get('currentDayProfitLossPercentage', 0),
            'marketValue': position.get('marketValue', 0),
            'maintenanceRequirement': position.get('maintenanceRequirement', 0),
            'longOpenProfitLoss': position.get('longOpenProfitLoss', 0),
            'previousSessionLongQuantity': position.get('previousSessionLongQuantity', 0),
            'currentDayCost': position.get('currentDayCost', 0),
            'cashAvailableForTrading': account_info['initialBalances'].get('cashAvailableForTrading', 0),
            'cashAvailableForWithdrawal': account_info['initialBalances'].get('cashAvailableForWithdrawal', 0),
            'cashBalance': account_info['initialBalances'].get('cashBalance', 0),
            'liquidationValue': account_info['initialBalances'].get('liquidationValue', 0),
            'longStockValue': account_info['initialBalances'].get('longStockValue', 0),
            'mutualFundValue': account_info['initialBalances'].get('mutualFundValue', 0),
            'accountValue': account_info['initialBalances'].get('accountValue', 0),
            'currentCashBalance': account_info['currentBalances'].get('cashBalance', 0),
            'currentLiquidationValue': account_info['currentBalances'].get('liquidationValue', 0),
            'longMarketValue': account_info['currentBalances'].get('longMarketValue', 0),
            'totalCash': account_info['currentBalances'].get('totalCash', 0),
            'currentAccountValue': account_info['currentBalances'].get('accountValue', 0),
            'aggregatedBalance': account['aggregatedBalance'].get('liquidationValue', 0)
        })

# Convert the flattened data into a pandas DataFrame
df = pd.DataFrame(flattened_data)

# Save the DataFrame to a CSV file
df.to_csv('account_data_with_positions.csv', index=False)

try:
    file_name = "account_data_with_positions.xlsx"
    
    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        # Write the DataFrame with the modified long quantities to Excel
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Access the workbook and worksheet after writing the DataFrame
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Create formats
        left_aligned_format = workbook.add_format({'align': 'left', 'bold': False})
        center_aligned_format = workbook.add_format({'align': 'center', 'bold': False})

        # Apply left alignment to the first column (account numbers)
        worksheet.set_column('A:A', 20, left_aligned_format)  # Column A for field names (account numbers)

        # Apply center alignment to columns B to Z
        worksheet.set_column('B:Z', 20, center_aligned_format)  # Center-align columns from B to Z

        # Adjust column widths based on the longest value in each column
        for col_num, column in enumerate(df.columns):
            max_len = max(df[column].astype(str).apply(len).max(), len(column))  # Get the max length of content or header
            worksheet.set_column(col_num, col_num, max_len + 2)  # Add some padding

        # Freeze the first 3 columns (A, B, C)
        worksheet.freeze_panes(0, 3)

    print(f"Excel file created successfully. {file_name}")
except Exception as e:
    print(f"Error creating Excel file: {e}")

# Create a new column for the current market price by dividing marketValue by longQuantity or shortQuantity
df['currentMarketPrice'] = df.apply(
    lambda row: row['marketValue'] / row['longQuantity'] if row['longQuantity'] > 0 else
               (row['marketValue'] / row['shortQuantity'] if row['shortQuantity'] > 0 else 0),
    axis=1
)

# Create a new column for the total cost by multiplying averagePrice by longQuantity or shortQuantity
df['totalCost'] = df.apply(
    lambda row: row['averagePrice'] * row['longQuantity'] if row['longQuantity'] > 0 else
               (row['averagePrice'] * row['shortQuantity'] if row['shortQuantity'] > 0 else 0),
    axis=1
)

# Create a new column for the change in price per share by subtracting averagePrice from currentMarketPrice
df['changePricePerShare'] = df.apply(
    lambda row: row['currentMarketPrice'] - row['averagePrice'] if row['longQuantity'] > 0 else
               (row['currentMarketPrice'] - row['averagePrice'] if row['shortQuantity'] > 0 else 0),
    axis=1
)

# Create an adjustment function based on the position description
def adjust_long_quantity(row):
    # Check if 'CD' is in the position description and adjust the quantity
    if 'CD' in row['positionDescription']:
        # row['longQuantity'] = row['longQuantity'] * 1000  # Modify the quantity by multiplying by 1000
        row['averagePrice'] = 1000  # Modify total cost by multiplying by changePricePerShare
        row['totalCost'] = row['longQuantity'] * row['averagePrice']  # Modify total cost by multiplying by changePricePerShare
    return row

# Apply the adjustment function to modify 'longQuantity' directly
df = df.apply(adjust_long_quantity, axis=1)

# Extract only the required columns for the smaller dataset
smaller_df = df[[
    'accountNumber', 
    'positionSymbol', 
    'positionDescription', 
    'positionType', 
    'longQuantity', 
    'shortQuantity', 
    'averagePrice', 
    'currentMarketPrice',
    'changePricePerShare',
    'totalCost',
    'marketValue', 
    'longOpenProfitLoss',
    'currentDayProfitLoss', 
    'currentDayProfitLossPercentage'
    ]]

# Renaming the columns for clarity if needed
smaller_df.columns = [
    'Account Number', 
    'Position Symbol', 
    'Position Description', 
    'Position Type', 
    'Long Quantity', 
    'Short Quantity', 
    'Cost / Share', 
    'Market Price / Share',
    'Change - Price / Share',
    'Total Cost',
    'Total Market Value',
    'Long Open Profit/Loss',
    'Current Day Profit/Loss', 
    'Current Day Profit/Loss (%)'
]

# Saving as an Excel file with appropriate formatting
try:
    file_name = "account_data_with_positions_v2.xlsx"
    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        # Write the flattened DataFrame to Excel
        smaller_df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Access the workbook and worksheet after writing the DataFrame
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Create formats
        left_aligned_format = workbook.add_format({'align': 'left', 'bold': False})
        center_aligned_format = workbook.add_format({'align': 'center', 'bold': False})

        # Red format for negative numbers
        red_format = workbook.add_format({'bg_color': '#FF0000', 'color': 'white', 'bold': False, 'align': 'center'})
        
        # Green format for positive numbers
        green_format = workbook.add_format({'bg_color': '#00FF00', 'color': 'black', 'bold': False, 'align': 'center'})

        # Apply left alignment to the first column (account numbers)
        worksheet.set_column('A:A', 20, left_aligned_format)  # Column A for field names (account numbers)

        # Apply center alignment to columns B to Z
        worksheet.set_column('B:Z', 20, center_aligned_format)  # Center-align columns from B to Z

        # Apply conditional formatting to specific columns if they exist in the DataFrame
        columns_to_format = ['Change - Price / Share', 'Long Open Profit/Loss', 'Current Day Profit/Loss', 'Current Day Profit/Loss (%)']

        # print(smaller_df.columns)  # Debugging: print the columns of the smaller DataFrame

        for col in columns_to_format:
          if col in smaller_df.columns:  # Ensure the column exists before applying formatting
              col_index = smaller_df.columns.get_loc(col) # Get the column index for formatting (1-based index for XlsxWriter)
              
              # Conditional formatting for negative values in red
              worksheet.conditional_format(1, col_index, len(smaller_df), col_index, 
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_format})

              # Conditional formatting for positive values in green
              worksheet.conditional_format(1, col_index, len(smaller_df), col_index, 
                                          {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_format})

        # Adjust column widths based on the longest value in each column
        for col_num, column in enumerate(df.columns):
            max_len = max(df[column].astype(str).apply(len).max(), len(column))  # Get the max length of content or header
            worksheet.set_column(col_num, col_num, max_len + 2)  # Add some padding

        # Freeze the first 3 columns (A, B, C)
        worksheet.freeze_panes(0, 3)

    print(f"Excel file created successfully. {file_name}")
except Exception as e:
    print(f"Error creating Excel file: {e}")
