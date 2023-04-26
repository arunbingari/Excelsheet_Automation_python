## Excel Sheet Automation using Openpyxl (Python)

# This is a Python automation script that works with excel sheets using the Openpyxl module. The script performs the following tasks:

1. Lists each company with the respective number of products in the inventory.
2. Calculates the total inventory value per supplier.
3. Lists all products with inventory less than 10.
4. Adds a new column to the excel sheet, which is the product of inventory and price.

## Requirements
This script requires the following:

1. Python 3.6 or higher
2. Openpyxl module

## Usage

1. Install the Openpyxl module by running the following command:
    pip install openpyxl
2. Save the script to your preferred directory.

3. Update the file path of the excel sheet on line 3 to match the location of your excel sheet.

4. Run the script by executing the following command in the terminal:
    python script_name.py

5. The output will be displayed in the terminal and saved in a new excel sheet with the name "inventory_with_total_value.xlsx".

## Note

Make sure the script and the excel sheet are in the same directory. If not, provide the absolute file path to the excel sheet on line 3.

## Conclusion

This script automates the process of calculating the total inventory value per supplier, listing each company with the respective number of products, and adding a new column to the excel sheet. It helps to save time and minimize errors that could occur during manual calculations.