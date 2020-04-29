This is a Python script that selects attributes/columns in a CSV/Excel file and creates an incremental index per each identifier.

Requirements: Python 3.7, xlrd library. Optional, but recommended: pandas library

Usage: python prepare_data.py [options] filename
Accepted formats: XLS(x) and CSV. In case the input is an Excel file(.xls, .xlsx), the script assumes it has only one sheet.
Options:
        --ids <ids>             Names of the columns that identify the staff in the dataset. IDs must be separated by commas ","
        --columns <columns>     Names of the columns to be shared from the entire dataset. Columns must be separated by commas ","
        --help                  Shows this help

Example: python prepare_data.py --ids "staff Number" --columns "Grade Type",Group,Reason,Start,End,"Total Duration" myfile.csv

Use --columns option to select the columns to be shared. Please make sure that any identification column must not be selected.
Use --ids option to select the columns that are used as staff identifiers (e.g., Staff Number). The script will create an incremental index for each staff.
