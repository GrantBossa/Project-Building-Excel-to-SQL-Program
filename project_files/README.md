# Excel-to-SQL
'''
The excelToCsv.py file is used for the following:

  a. Traverse a given file directory to locate XLSX files.

  b. Process each file and worksheet (Clean up data inconsistencies).
  
     i. Standardize the data format.
     
     ii. Organize data by date, time, machine_ID. etc.
  
  c. Create a single CSV formated file.

A successfully executed program should leave behind one CSV file in the directory
the program was executed in, with the name supplied in the arguments.

Errors will be executed for the following:

  a. Not using a .csv extension for the output file.
  
  b. Using an directory to traverse that has no .xlsx files.

A mock-filesystem is provided in order to see how this program functions.

To use this directory, in the command below 
substitute directory_to_traverse with mock-filesystem 

The correct commandline usage is:
  excelToCsv.py output.csv directory_to_traverse/
