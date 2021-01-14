# URL for website guide = https://medium.com/swlh/how-to-move-data-from-one-excel-file-to-another-using-python-a513491234f4
# Author - Stefan G. Creadore
print("Author - Stefan G. Creadore")

# Usage Agreement - Copyright 2021 Stefan Creadore - Free to use or modify as long as credit is given in any public usage or publications.
print("Usage Agreement - Copyright 2021 Stefan Creadore - Free to use or modify as long as credit is given in any public usage or publications.")


# This package is for transforming and transferring sample requisition data into the Ovation bulk upload requisition template.
print("This package is for transforming and transferring sample requisition data into the Ovation bulk upload requisition template.")


# If you have questions or concerns contact me through LinkedIN - https://www.linkedin.com/in/stefan-creadore/
print("If you have questions or concerns contact me through LinkedIN - https://www.linkedin.com/in/stefan-creadore/")

# I recommend using a virtual environment for this Python script.
# begin by installing pandas by using python pip install pandas or python pip --install pandas.
# After installing pandas use "import pandas as pd" to import the pandas package.
# Install the package dependency "xlrd" to read in excel files in file.xls format only by using pip. Example: pip install xlrd
# Install openpyxl by using "pip install openpyxl" if you want to import xlsx files or other types of excel file formats.

import pandas as pd
import time
from playsound import playsound

#playsound('initializing.mp3')

# Now we will create three variables. 1) The spreadsheet_file will be the source file workbook. 2) worksheets variable will be the sheets in that source file. 3) appended_data as an empty array.

spreadsheet_file = pd.ExcelFile('samples.xlsx')
worksheets = spreadsheet_file.sheet_names
appended_data = []

print("Phase 1 will begin now... Loading Dataframes from excel file.")
time.sleep(3)
# Next we use a "for" loop to get the data we want from each sheet in the source data file. This approach is highly scalable so no matter how many sheets you have within an Excel file the script will always check each sheet. Note that df = dataframe, identifier is the sample identifiable ID/name.
# I will use the value of 0 for the header row since our header row in the sheet is Excel row 3 and the row values recognized by Pandas are on a zero-based index (meaning row 1 in Excel would be a header = 0 value).

for sheet_name in worksheets:
    identifier = 'Identifier'
    Date_Received = 'Date Received'
    date = 'Date Received'
    df = pd.read_excel(spreadsheet_file, sheet_name, header=0)
    df = df[['Date Received', identifier]].where(df[date]>"8/5/2019") # This shows the Sample Identifier/Sample Name in relation to the Date it was Received in your lab or date of your choosing. You can change the date in this line at the >"11/25/2019" part. You can make it greater than (>) or less than (<) or equal to (=).
    df = df.dropna() # Removes the NaN values in the table.
    df['Hospital Location'] = sheet_name # Creates the location column and adds it to the dataframes.
    df = df[['Hospital Location', 'Identifier', date]]
    appended_data.append(df) # Adds each dataframe to an array.

#playsound('process_completed.mp3')
time.sleep(3)

print("Now Beginning Phase 2")


appended_data = pd.concat(appended_data) # Joins all the dataframes in the array into a single dataframe.
print(appended_data)
appended_data.to_excel('C:/Users/SCreadore/Desktop/Ovatioin-requisition-automation/requisition2.xlsx')

print("Task Execution Status = Completed. Congratulations the dataframes now contain every column in the sheet.")

