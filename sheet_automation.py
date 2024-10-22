from openpyxl import Workbook
from win32com.client import Dispatch

workbook = Workbook()
sheet = workbook.active

#Create Sheets
sheet["A1"] = "Hello"
sheet["B1"] = "World"

#Save your file
workbook.save(filename="example.xlsx")

#Change sheet value
cell = sheet["A1"]

cell.value = "Goodbye"

#Open your file with Excel
#If you have an error, you can change the path to your archive
#For example: C:/Users/YourName/Desktop/example.xlsx
#Don't forget to change the extension to .xlsx

#Open Excel
#x1 = Dispatch("Excel.Application")
#x1.Visible = True

#wb = x1.Workbooks.Open("your archive path")