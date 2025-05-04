# Excel-to-SQL

##  The correct commandline usage is:
###  excelToSQL.py output.db dirOfExcelFiles

## The excelToSQL.py file is used for the following:

  a. Traverse a given file directory to locate XLSX files.

  b. Process each file and worksheet (Clean up data inconsistencies).
  
     i. Standardize the data format.
     
     ii. Organize data by date, time, machine_ID. etc.
  
  c. Create a single sql formated database table in a SQL database.

## ----------------------------------------------
A successfully executed program should leave behind one db file in the directory
the program was executed in, with the name supplied in the arguments. 
If this file already exists, it will be overwritten.

Note: 
Be advised that audit trail fields have been added to the output file in order to verify files processed. 
I used them to verify complete processing of the Excel files.
These can be deleted or commented out based upon your need.


## ----------------------------------------------
Errors will be executed for the following:

  a. Not using the correct commandline format.
  
  b. Not using a .db extension for the output file.
  
  c. Using an directory to traverse that has no .xlsx files.

  d. Failing to convert files, for any reason, before writing the output file.

## ----------------------------------------------
mock-filesystem and empty-filesystem directories are provided in order to see how this program functions.

To use these directories, in the command above 
substitute dirOfExcelFiles with either mock-filesystem or empty-filesystem.


# Excel-to-CSV

## The correct commandline usage is:
###  excelToCSV.py output.csv dirOfExcelFiles

## The excelToCSV.py file is used for the following:

  a. Traverse a given file directory to locate XLSX files.

  b. Process each file and worksheet (Clean up data inconsistencies).
  
     i. Standardize the data format.
     
     ii. Organize data by date, time, machine_ID. etc.
  
  c. Create a single CSV formated file.

## ----------------------------------------------
A successfully executed program should leave behind one CSV file in the directory
the program was executed in, with the name supplied in the arguments. 
If this file already exists, it will be overwritten.

Note: 
Be advised that audit trail fields have been added to the output file in order to verify files processed. I used them to verify complete processing of the Excel files.
These can be deleted or commented out based upon your need.


## ----------------------------------------------
Errors will be executed for the following:

  a. Not using the correct commandline format.
  
  b. Not using a .csv extension for the output file.
  
  c. Using an directory to traverse that has no .xlsx files.

  d. Failing to convert files, for any reason, before writing the output file.

## ----------------------------------------------
mock-filesystem and empty-filesystem directories are provided in order to see how this program functions.

To use these directories, in the command above 
substitute dirOfExcelFiles with either mock-filesystem or empty-filesystem.
