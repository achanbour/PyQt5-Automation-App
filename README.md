# PyQt5-Automation-App

The KRI automation app described below is a single page GUI (Graphical User Interface) automating three core tasks in the KRI (Key Risk Indicator)collection process:

1. Generating Excel files from the KRI master table of the operational risk database
2. Sending emails to a list of recipients stored in the operational risk database attaching the generated files from step 1
3. Updating the KRI database with any responses received to the emails sent in step 2

The automation of each of the above three tasks is described in its corresponding ```.py``` file:

1. ```generate_excel.py``` defines the following methods to generate the KRI Excel reports which are then sent to the recipients in step 2.
   1. ```generate_kri_report()``` is a function performing the following steps:
      1. Reads the KRI master table as a pandas dataframe
      2. Applies required filters to the KRI dataframe to get a filtered KRI dataframe
      3. Generates an Excel file from the filtered KRI dataframe
      4. Archives the Excel file from the previous step in a specified folder
   2. Various helper methods are defined to make sure the folders in which the Excel reports are to be stored exist, and if not, create them during execution. The said helper methods are called in the ```generate_kri_report()```  method.
   3. ```xl_stylisation()``` is a helper method called during the execution of ```generate_kri_report()```. It opens the Excel report generated by ```generate_kri_report()``` and applies various formatting adjustments such as resizing columns and rows based on cell content, and applying color formatting rules to the cells.

2. ```send_emails.py``` defines a single method ```send_automatic_emails()``` for sending automatic emails. This method executes the following steps in order:
   1. Reads the mailing list table as a pandas dataframe
   2. Iterates over the Excel reports generated in step 1. For each report:
      1. Filters the mailing list dataframe to identify the corresponding recipient(s)
      2. Accesses the email server via the local Outlook app using the ```pywin32``` package and creates a new email message with the Excel report as attachment
      3. Sends the email to the recipient(s) listed in the filtered mailing list dataframe
      4. Archives the email from the previous step in a specified folder

3. ```update_db.py``` defines a single method ```update_kri_table()``` for updating the KRI table with the results obtained in the reports returned to the automated emails sent in step 2. This step is executed only after responses have been received and archived in the dedicated folders which is accomplished automatically using an [inbox parser](https://github.com/achanbour/Windows-Outlook-inbox-parser). The ```update_kri_table()``` method executes the following steps in order:
   1. Reads each archived report as a pandas dataframe and stores the KRI results in a temporary storage location (Python dictionary)
   2. Reads the KRI table from the database as a pandas dataframe
   3. Writes the KRI results to the KRI table

The implementation of each task is further detailed in its respective Python file.
