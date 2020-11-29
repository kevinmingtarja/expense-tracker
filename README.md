# expense-tracker
Script that I made to track expenses using python, Gmail API, Google Drive API, and Google Sheets API.
This script uses Gmail API to read the transaction email notifications and extract the date, description, and amount of each transaction, which then is returned as a pandas dataframe object and then gets passed to a Google Sheets spreadsheet using the Google Drive and Sheet API.
