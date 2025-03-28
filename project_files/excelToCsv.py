# 
# excelToCsv.py
# Grant Bossa
# SDEV 1100
# 2/27/2025
#

import sys              # command line arguments
import os               # file and directory navigation
import pandas as pd     # Excel file manipulation

# Initalize Variables
all_data = []                       # For creating csv file

def traverse_and_convert(directory):
    countOfExcelFilesFound = 0      # Verify number of files found
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
        # Read Excel file and get all worksheets
        excel_data = pd.read_excel(file_path, sheet_name=None)

        # for each worksheet in the file
        for sheet_name, df in excel_data.items():
            # Add metadata columns
            df['From File'] = file_path             # save File path 
            df['From Worksheet'] = sheet_name       # Save Worksheet name
            df['From Row Number'] = df.index + 1    # Save Row count
            df['From Col Count'] = df.shape[1]-3    # Save Column count (minus 3 added columns)

            all_data.append(df)                     # Add records to all_data 

    except Exception as e:
        print(f"Failed to convert {file_path}: {e}")

def main():
    try :
        # Verify correct program arguments are being used
        if len(sys.argv) != 3:
            print(f"Usage: {sys.argv[0]} output.csv dirOfExcelFiles/")
            sys.exit(1)  # Exit with an error code - incorrect arguments for program to run
        
        if ((type(sys.argv[1]) == str) 
            and (sys.argv[1].lower().endswith(".csv"))):
            csv_file = sys.argv[1]              # Set csv filename
        else:
            print(f"The output file should be a .csv filename! : Usage: {sys.argv[0]} output.csv dirOfExcelFiles/")
            sys.exit(1)  # Exit with an error code - incorrect arguments for program to run

        # Set Directory to traverse for Excel files to convert
        traverse_directory = sys.argv[2]    

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
        column_order = ['Date', 'Time', 'Machine ID', 'Sensor 1', 'Sensor 2', 'Temperature', 'Pressure', 
                        'Flow Rate', 'Status', 'Employee', 'From File', 'From Worksheet', 'From Row Number', 'From Col Count']
        combined_df = combined_df[column_order]

        # Write the completed file
        combined_df.to_csv(csv_file, index=False)

        # If no errors. Print file created message
        print(f"CSV file {csv_file} created.")

        ''' COPIED CODE FROM VIDEO, IT NEEDS FLESHED OUT!!!!!
        # Connect to SQL database
        out_db_name = output_name + ".db"
        table_name = "mine_data"
        con = (sqlite).connect(out_db_name)
        rows_stored = df.to_sql(table_name, con, if_exists = "fail", index=True)
        con.close()
        print(f"{out_db_name) successfully created with [rows_stored} rows")
        '''

    except KeyboardInterrupt:
        print("Program interrupted by user.")
    except Exception as e:
        print(f" {Exception}: {e}")

if __name__ == '__main__' :
    main()
    
