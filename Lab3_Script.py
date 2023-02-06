
from sys import argv
import os
import sys
from datetime import date
import pandas as pd
import re
def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    try:
        sales_csv = argv[1]
    except:
        print("Error: Please provide a file path")
        sys.exit()
    # Check whether provide parameter is valid path of file
    if os.path.isfile(sales_csv):
        print(str(os.path.isfile(sales_csv)))
        return sales_csv
    else:
        print("Error: File path doesn't exist")
        sys.exit()

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    directory_path = os.path.dirname(sales_csv)
    # Determine the name and path of the directory to hold the order data files
    order_dir_name = 'Orders_'
    current_date = date.today().isoformat()
    order_dir_name += current_date
    order_dir_path = os.path.join(directory_path, order_dir_name)
    # Create the order directory if it does not already exist
    if os.path.exists(order_dir_path) is False:
        os.makedirs(order_dir_path)
    return order_dir_path

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    
    # Import the sales data from the CSV file into a DataFrame
    sales_df = pd.read_csv(sales_csv)
    
    # Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])
    
    # Remove columns from the DataFrame that are not needed
    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)
    
    # Group the rows in the DataFrame by order ID
    groups = sales_df.groupby('ORDER ID')
   
    # For each order ID:
    for order_id, order_df in groups:
        
        # Remove the "ORDER ID" column
        del order_df['ORDER ID']
        
        # Sort the items by item number
        order_df.sort_values(by=['ITEM NUMBER'], inplace=True)
        
        # Append a "GRAND TOTAL" row
        grand_total = order_df['TOTAL PRICE'].sum()
        GRAND_TOTAL_df = pd.DataFrame({'ITEM PRICE':['GRAND TOTAL'], 
                                    'TOTAL PRICE':[grand_total]})
        order_df = pd.concat([order_df, GRAND_TOTAL_df])
        print(order_df)
        
        # Determine the file name and full path of the Excel sheet
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W', '', customer_name)
        order_file_name = f'Order{order_id}_{customer_name}.xlsx'
        order_file_path = os.path.join(orders_dir, order_file_name)
        
        # Export the data to an Excel sheet
        sheet_name_to_use = f'Order #{order_id}'        
        order_df.to_excel(order_file_path, index=False, sheet_name = sheet_name_to_use)
        # TODO: Format the Excel sheet
    pass

if __name__ == '__main__':
    main()