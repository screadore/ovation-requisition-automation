# ovation-requisition-automation
# Disclaimer I do not claim any ownership or rights to Ovation. I am just facilitating the automation of their platform for open source usage.
This script uses the Pandas Python package to automate the input of sample information from one excel file into the requisition template for bulk uploading.

How to use:
1) Download the repo to your desktop.

2) Open your CMD/Terminal/GitBash, whichever you use. Alternatively PyCharm is a useful tool.

3) cd Path/User/Desktop/Ovatioin

4) Install pandas if you don't currently have it installed using - "pip install pandas" for a virtual environment or to install to system pip install --user pandas.

5) Then install playsound using "pip install playsound" for a virtual environment or "pip install --user playsound" for a system install.

6) Then copy and paste your input excel file into the ovatioin folder.

7) Set the PATH to your excel file in the "ovation.py" script on line 30. It looks like this "spreadsheet_file = pd.ExcelFile('samples.xlsx')"
  - Make sure this is in the .xlsx format or the script will NOT WORK!!!!

8) You can select the date that you want to pull samples from on line 44 of the "ovation.py" script it should say "df = df[['Date Received', identifier]].where(df[date]>"8/5/2019")"
  - Just change the date "8/5/2019" nothing else. 
  - Please note that this will make it so that samples on the date selected and newer will be pulled into the dataframe array.
  
9) To change the name of your output file go to line 58 in the "ovation.py" script, which will look something like this "appended_data.to_excel('C:/Users/SCreadore/Desktop/Ovatioin-requisition-automation/requisition2.xlsx')"
    - Then change the file name there.
    - Note if you do not change the file name the output file name will be "requisition2.xlsx"
    - Also Make sure you set the path to where your ovatioin folder is.

10) run the ovation.py script and you should be good to go.

# Please note this is still a version 1 beta of the script. Will be updating over time to make it more efficient.
