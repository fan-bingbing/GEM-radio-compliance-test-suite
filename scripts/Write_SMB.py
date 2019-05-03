import visa
from openpyxl import load_workbook

rm = visa.ResourceManager()
SMB = rm.open_resource('USB0::0x0AAD::0x0054::106409::INSTR') # SMB100A

SIGEN_file_write = load_workbook(filename = "SIGEN_Write.xlsx") # create a workbook from existing .xlsx file
sheet = SIGEN_file_write["Setup"] # load setup sheet in .xlsx to sheet
frequency = sheet["B1"].value # get frequency value from sheet
level = sheet["B2"].value # get level value from sheet
RF_output = sheet["B3"].value # get RF output status from sheet

print (f"Set the SigGen frequency to {frequency} MHz.")
print (f"Set the SigGen level to {level} dBm.")
print (f"Set the SigGen RF output ON.")

SMB.write(f":FREQ {frequency}MHz ") # write frequency value to SMB100A
SMB.write(f":POW {level}dBm ") # write level value to SMB100A
SMB.write(f":OUTP {RF_output} ") # write RF output status to SMB100A
