import win32com.client


#def printToPdf():
# Initialize the Excel application
excel_app = win32com.client.Dispatch("Excel.Application")

# Open the workbook
workbook_path = r"C:\Users\borko.kovacevic\Desktop\Project\nalozi-štampa.xlsm"  # Change this to your workbook's path
workbook = excel_app.Workbooks.Open(workbook_path)

# Make Excel visible (optional)
excel_app.Visible = True

# Access the sheet with the button
sheet = workbook.Sheets("NalogKT (2)")  # Change "Sheet1" to your sheet name


# Press the button by calling the macro associated with it
# Assuming the macro is named "MyMacro"
excel_app.Application.Run("KTCreatePDF")

# Save and close the workbook
workbook.Save()
workbook.Close()

# Quit the Excel application
excel_app.Quit()
