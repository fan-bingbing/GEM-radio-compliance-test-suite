from openpyxl import load_workbook
limit_line_factor = 0 # limit line indicator for conducted spurious test
RFile_write = load_workbook(filename = "Test_Result.xlsx") # load Test_Result.xlsx
RSheet1 = RFile_write["Cond_Spur"] # load "Ferror_Pow" sheet in .xlsx
