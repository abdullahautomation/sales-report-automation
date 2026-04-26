# Sales Report Automation Tool

## What it does
This tool takes a messy Excel sales file and automatically:
- Cleans the data (missing values, inconsistent names)
- Generates a professional formatted report
- Creates Monthly and Product summary sheets
- Adds bar charts automatically

## Requirements
- Python 3.x
- Libraries: pandas, openpyxl

## Installation
pip install pandas openpyxl

## How to use
1. Place your Excel file in the same folder as the script
2. Double click on gui.py and Run from python. (If python is not Installed in your computer then install it first because it is compulsory.)
3. Enter your input file name: raw_sales_data.xlsx
4. Enter your output file name: my_report.xlsx

## Input file requirements
Your Excel file must have these columns:
- Order_ID
- Salesperson
- Product
- Region
- Sale_Amount
- Month

## Output
A professionally formatted Excel file with:
- Cleaned Data sheet
- Monthly Summary sheet with chart
- Product Summary sheet with chart