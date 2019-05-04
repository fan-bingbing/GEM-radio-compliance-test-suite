import os

def Screenshot(filename)
print("Taking instrument screenshot and saving it to the PC... ")
FSV.write("SYST:DISP:UPD ON")
FSV.write("DISP:TRACE1:MODE VIEW")# make sure plots are stable
FSV.write("HCOP:DEV:LANG BMP")
FSV.write("HCOP:CMAP:DEF4")
FSV.write(f"MMEM:NAME \'c:\\temp\\Dev_Screenshot.bmp\'")
# keep overwriting Dev_Screenshot.bmp in FSV C drive, note the usage of backslash \

FSV.write("HCOP:IMM")
# take the screenshot now
FSV.query("*OPC?")  # Wait for the screenshot to be saved

file_data = FSV.query_binary_values(f"MMEM:DATA? \'c:\\temp\\Dev_Screenshot.bmp\'", datatype='s',)[0]
new_file = open(f"c:\\Temp\\{file_name}", "wb")# extract file_name string using (f"{}")
new_file.write(file_data)
new_file.close()
print(f"saved to PC c:\\Temp\\{file_name}\n") # extract file_name string using (f"{}")
