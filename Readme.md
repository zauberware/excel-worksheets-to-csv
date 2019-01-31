# Export Worksheets from Excel to CSV

This script will help you to extract Worksheets from an Excel file. It already saved me days of my lifetime!

## Run the example

Clone this repository and try to run the example file. The file has 3 worksheets: `Worksheet 1`, `Another data set` and `Last`
 
1. `git clone git@github.com:zauberware/excel-worksheets-to-csv.git && cd excel-worksheets-to-csv`
2. `python worksheet_to_csv.py example.xlsx`

Now you should see `Worksheet 1.csv`, `Another data set.csv` and `Last.csv` in the exports folder.


## Export all Worksheets

To export all worksheets to CSV run `python worksheet_to_csv.py example.xlsx` like in the example.


## Export specific Worksheets

To export only specific worksheets you can specify them comma seperated as the second argument of the script.

**Example: Export only `Worksheet 1` and `Last`:**

`python worksheet_to_csv.py example.xlsx "Worksheet 1,Last"`


## Prerequisites

* Be sure `openpyxl` and `imp` is installed.
* Convert your file into XLSX format if needed. It not works with other excel formats.
