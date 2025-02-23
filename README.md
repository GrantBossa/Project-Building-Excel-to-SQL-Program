# Excel-to-SQL
# Sdev 1100 files for converting Excel files to CSV
#
# Project Building Excel to SQL Program
# Project - Building Stage
# 1. Read and clean data from Excel files to CSV format
#    a. Traverse file directory to locate XLSX files
#    b. Process each file and worksheet (Clean up data inconsistencies)
#        i. Standardize the data format
#        ii. Organize data by date, time, machine_ID
#    c. Create a CSV file formated to convert into SQL database
# 3. Convert CSV formated data into an SQL database
# 4. Build the GUI for accessing the data online.

'''
Project Pt. 1 Requirements
SDEV 1100 SP 2025
Due on February 28th, 2025
__________________________________________________________
2/4/2025 Update:
• I mention to “trust that trust that Date, Time, MachineID, Sensor1, etc will be in
every file we’re working with”. What I meant is that the fields will have the same
name from file to file, so date will always be named ‘Date’, and time will always be
‘Time’. However, some data may not have some fields that other fields have. For
now, assume that all data will have a date and time. Pandas should be able to handle
missing fields, and will automatically write them as blank in the csv file.

• If there are two entries that have the same date and time, they should be right next to
each other in the csv file, and sorted by the machineId.

• When I say that the program should “complain and exit” that doesn’t mean it should
dump out a bunch of debugging information like it does by default. Your program
should try to do something, and exit with a simple, clean, error message if it cannot
do something. This is called having your program gracefully fail.

Until now, most of the planning has been up to you. Beginning a large project can be
uncomfortable, especially when the project is for a customer and it’s your first time doing
anything like it. This uncomfortable feeling is why they pay us the big bucks. We don’t do this
because it’s easy, but because we thought it would be easy. This document should help you
understand in more technical detail than the BRD what is expected out your program by the first
submission. You may need to read it multiple times. More importantly, you may even need to
ASK QUESTIONS. The first submission of this project is worth 20% of your grade in this class.
This program will work on the supplied mock-filesystem in a GitHub Codespaces
environment. Why Codespaces? Because it is a containerized Linux environment (Ubuntu) and
allows us to have access to the same development environment to run our code on. No “but it
works on my machine” allowed. Your python script needs to work on a Linux filesystem, not
Windows. I will never run your program on Windows. If you choose to develop your script
outside of Codespaces, make sure that it still works in Codespaces before turning it in because
that’s where I will run it when it comes time to grade.

Ultimately, we’re going to be tying our Python script into a backend framework and
database, but the goal for the first part of the semester is to get all the data from the Excel files
into one all-encompassing CSV file.

When it comes time to test your program, one of your tests
should be checking that every row and column from the Excel files has made its way to your
CSV file. Not losing data is top priority. Make sure no files are being skipped and no data is
being lost from each file. During development, I recommend creating and working on a
manageable set of short files based on the provided mock files. That way, it’s easier to visually
check if everything is working. I will be testing your script on the mock filesystem I’m providing
to you.
Your program only needs to be runnable from the command line. The user should be able
to run the program by typing python excelToCsv.py myOutput.csv mock-filesystem/ (It
doesn’t actually matter what you name your program / directory; just as long as I can tell which
is which). This command will run the program in its entirety. Use a “main” function to handle
arguments and kick things off. Accepted arguments are a string for the output filename, and a
non-empty directory. If the argument is not a valid path to a directory, the program should
complain and exit. If recursively searching the directory doesn’t yield any xlsx files, complain
and exit. Remember: in UNIX like filesystems, ‘./’ is the current directory and ‘../’ is the
parent directory.
You’re going to have to figure out how to traverse the file system and find .xlsx files. For
every .xlsx file you find, use pandas.read_excel() to read the data into a DataFrame. For this
first part of the project sort the data by date and time, oldest to newest. If two entries have the
same time, sort by the Machine ID.
The fields that the files have right now can be used in your code to find data. So trust that
Date, Time, MachineID, Sensor1, etc will be in every file we’re working with. That way you can
reference these fields as you use your DataFrame. This might not be the case in the real world,
but perhaps we can add some AI functionality in the future.
A successfully executed program should only leave behind one CSV file in the directory
the program was executed in, with the name supplied in the arguments. Do NOT manipulate the
Excel files in any way. If the program encounters errors and needs to bail out, don’t write an
unfinished output.csv unless the error has to do with writing the file itself. Error handling is a
must, and you should try adding some broken/badly formatted CSV files to test on. The program
should run successfully on the mock-filesystem, but that doesn’t mean I won’t try to break your
program with other data. We’ll talk more about this when we’re testing in week 6.
More miscellaneous notes:
- If a cell is blank, leave it as an empty string in the csv file.
- Make sure that all floats are rounded to one decimal place.
- The final output.csv should have all fields in the same order. If the file doesn’t contain
the fields, leave the would-be value in that field blank. (Yes, the ‘test.xlsx’ file will be
included in the final output).
'''