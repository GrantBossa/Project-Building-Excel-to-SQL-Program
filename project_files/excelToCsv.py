# 
# Grant Bossa
# SDEV 1100
# 2/23/2025
# excelToCsv.py
#

import sys              # command line arguments
import os               # file and directory navigation
import pandas as pd     # Excel file manipulation

# Initalize Variables

all_data = []                       # For creating csv file

def traverse_and_convert(directory):
    countOfExcelFilesFound = 0
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                countOfExcelFilesFound += 1
                file_path = os.path.join(root, file)
                convert_to_csv_format(file_path)
    # If there were no files found processing
    if countOfExcelFilesFound == 0 :
        print(f"Error: No Excel Files found in the directory '{directory}' ")
        sys.exit(1)  # Exit with an error code - incorrect arguments for program to run

def convert_to_csv_format(file_path): 
    # Used to create the dataframe before converting to csv file
    try:
        excel_data = pd.read_excel(file_path, sheet_name=None)

        # for each worksheet in the file
        for sheet_name, df in excel_data.items():

            # Add metadata columns
            df['From file'] = file_path
            df['From Worksheet'] = sheet_name
            df['From Row number'] = df.index + 1

            all_data.append(df)

        combined_df = pd.concat(all_data)

    except Exception as e:
        print(f"Failed to convert {file_path}: {e}")

# Control the column format of appended dataframes
def sort_columns(df, column_order):
    return df[column_order]

def main():
    try :
        # Verify correct program arguments are being used
        if len(sys.argv) != 3:
            print(f"Usage: {sys.argv[0]} output.csv dirOfExcelfiles/")
            sys.exit(1)  # Exit with an error code - incorrect arguments for program to run
        
        if ((type(sys.argv[1]) == str) 
            and (sys.argv[1].lower().endswith(".csv"))):
            csv_file = sys.argv[1]              # Set csv filename
        else:
            print(f"The output file should be a .csv filename! : Usage: {sys.argv[0]} output.csv dirOfExcelfiles/")
            sys.exit(1)  # Exit with an error code - incorrect arguments for program to run

        # Set Directory to traverse for Excel files to convert
        traverse_directory = sys.argv[2]    
        total_row_count_of_datafile = 0     # Acumulator for verifying row counts added

        # verify traverse path is valid
        if (os.path.isdir(traverse_directory)):
            traverse_and_convert(traverse_directory)
        else:
            print(f"Error: Directory {traverse_directory} does not exist!")
            sys.exit(1)  # Exit with an error code - incorrect arguments for program to run


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
        column_order = ['From file', 'From Worksheet', 'From Row number', 'Date', 'Time', 
                        'Machine ID', 'Sensor 1', 'Sensor 2', 'Temperature', 'Pressure', 
                        'Flow Rate', 'Status', 'Employee']
        sorted_df = sort_columns(combined_df, column_order)

        # Write the completed file
        sorted_df.to_csv(csv_file, index=False)

        #if no errors. Print file created message
        print(f"CSV file {csv_file} created.")

    except KeyboardInterrupt:
        print("Program interrupted by user.")
    except Exception as e:
        print(f" {Exception}: {e}")

    
    

if __name__ == '__main__' :
    main()

