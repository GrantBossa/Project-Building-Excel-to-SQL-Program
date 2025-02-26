# 
# Grant Bossa
# SDEV 1100
# 2/23/2025
# excelToCsv.py
#

import sys
import os
import pandas as pd

# Initalize Variables
# Set up for creating csv file
all_data = []



def traverse_and_convert(directory):
    countOfExcelFilesFound = 0
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                countOfExcelFilesFound += 1
                file_path = os.path.join(root, file)
                convert_to_csv(file_path)
    if countOfExcelFilesFound == 0 :
        print("Error: No Excel Files found!")

def convert_to_csv(file_path):
    '''
    # Example usage excel_to_csv
        excel_file_path = 'your_excel_file.xlsx'
        csv_file_path = 'combined_output.csv'
        # You can use path and filename as arguments
            excel_to_csv(excel_file_path, csv_file_path)
        # Or you can use just the filename
            excel_to_csv(excel_file, csv_file):
    '''
    try:
        excel_data = pd.read_excel(file_path, sheet_name=None)
        #csv_file_path = file_path.rsplit('.', 1)[0] + '.csv'
        #excel_data.to_csv(csv_file_path, index=False)
        """
        Imports all worksheets from an Excel file into a single CSV file.

        Args:
            excel_file (str): Path to the Excel file.
            csv_file (str): Path to the output CSV file.
        """
        #excel = pd.ExcelFile(file_path)

        #for sheet_name in excel.sheet_names:
        for sheet_name, df in excel_data.items():
            #df = excel.parse(sheet_name)
            # Add metadata columns
            df['from file'] = file_path
            df['worksheet'] = sheet_name
            df['row number'] = df.index + 1
            #sorted_df = sort_columns(df, column_order)
            all_data.append(df)
            total_row_count_of_datafile += df.shape[0] # row count of datafile
            print(f"Appending {file_path} worksheet {sheet_name}")

        combined_df = pd.concat(all_data)

        #print(f"Converted {file_path} to {csv_file_path}")
        print(f"Converted {file_path}")
    except Exception as e:
        print(f"Failed to convert {file_path}: {e}")

# Control the column format of appended dataframes
def sort_columns(df, column_order):
    return df[column_order]

def main():

    # Verify correct program entries
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} output.csv dirOfExcelfiles/")
        sys.exit(1)  # Exit with an error code - incorrect arguments for program to run

    csv_file = sys.argv[1]              # Set csv filename
    traverse_directory = sys.argv[2]    # Set Directory to traverse for Excel files to convert

    total_row_count_of_datafile = 0
    # verify traverse path is valid
    if (os.path.isdir(traverse_directory)):
        traverse_and_convert(traverse_directory)
    else:
        print(f"Error: Directory {traverse_directory} does not exist!")
        return

    # Prepare to write the datafile to csv file
    # Create combined dataframe
    combined_df = pd.concat(all_data)
    # Sort before converting
    combined_df = combined_df.sort_values(['Date','Time', 'Machine ID'])
    # Change float values to 1 decimal place
    combined_df['Sensor 1'] = combined_df['Sensor 1'].round(1)
    combined_df['Sensor 2'] = combined_df['Sensor 2'].round(1)  
    combined_df['Temperature'] = combined_df['Temperature'].round(1)  
    combined_df['Pressure'] = combined_df['Pressure'].round(1)  
  
    # Specify the desired column order and sort
    column_order = ['From file', 'Worksheet', 'Row number', 'Date', 'Time', 'Machine ID', 'Sensor 1', 'Sensor 2', 'Temperature', 'Pressure', 'Flow Rate', 'Status', 'Employee']
    sorted_df = sort_columns(combined_df, column_order)

    print(f"total rows added: {total_row_count_of_datafile}")
    print(f"total rows of df: {sorted_df.shape[0]}")
          
    # Write the completed file
    sorted_df.to_csv(csv_file, index=False)
    

    
    

if __name__ == '__main__' :
    main()
# 1. Week5Reading.pdf https://westernwyoming.instructure.com/courses/17647/files/2687613?wrap=1
# 2. https://pandas.pydata.org/docs/reference/io.htmlLinks to an external site.
