import openpyxl as xl


wb = xl.load_workbook('C:\\Resources\\Program_Xl_Sheets\\SampleWorkBook.xlsx')
print(wb.sheetnames)                                               # Get all the sheet names
sheet = wb['First_Sheet']                                          # Get sheet
# sheet = wb.get_sheet_by_name('First_Sheet')                      # deprecated
sheet.title = "First1_Sheet"                                       # Rename sheet by changing its title
# changes will not reflect until we save the workbook
wb.save('C:\\Resources\\Program_Xl_Sheets\\SampleWorkBook.xlsx')   # Save workbook
