# 
# Grant Bossa
# SDEV 1100
# 2/23/2025
# excelToCsv.py
#

import sys
import pandas as pd

def excel_to_csv(excel_file, csv_file):
    """
    Imports all worksheets from an Excel file into a single CSV file.

    Args:
        excel_file (str): Path to the Excel file.
        csv_file (str): Path to the output CSV file.
    """
    excel = pd.ExcelFile(excel_file)
    all_data = []
    for sheet_name in excel.sheet_names:
        df = excel.parse(sheet_name)
        all_data.append(df)

    combined_df = pd.concat(all_data)
    combined_df.to_csv(csv_file, index=False)



def main():

  if len(sys.argv) != 3:
    print(f"Usage {sys.argv[0]} output.csv dirOfExcelfiles/")
    sys.exit(1)  # Exit with an error code

  csv_file = sys.argv[1]    # csv filename
  traverse_directory = sys.argv[2]    # Directory to traverse for Excel files to convert

  print(f"Creating CSV file: {csv_file}")
  print(f"Scanning {traverse_directory} for Excel files to convert")
  
  # Example usage:
  excel_file_path = 'your_excel_file.xlsx'
  csv_file_path = 'combined_output.csv'
  excel_to_csv(excel_file_path, csv_file_path)
  excel_to_csv(excel_file, csv_file):

if __name__ == '__main__' :
  main()
# 1. Week5Reading.pdf https://westernwyoming.instructure.com/courses/17647/files/2687613?wrap=1
# 2. https://pandas.pydata.org/docs/reference/io.htmlLinks to an external site.