import visa
import time
from openpyxl import load_workbook

rm = visa.ResourceManager()
FSV = rm.open_resource('TCPIP0::192.168.10.9::hislip0::INSTR') # Spec An

FSV_file_write = load_workbook(filename = "Test_Setup.xlsx") # create a workbook from existing .xlsx file
sheet = FSV_file_write["Analog_Demod"] # load setup sheet in .xlsx to sheet

Centre_frequency = sheet["B1"].value #
Dev_PerDivision = sheet["B2"].value #
Demod_BW = sheet["B3"].value #
AF_Couple = sheet["B4"].value #
RF_level = sheet["B5"].value #
Attenuation = sheet["B6"].value # get attenuation
RefLev_offset = sheet["B7"].value# get RFlevel offset
Trace_Peak = sheet["B8"].value # get trace to Pos peak
Demod_MT = sheet["B9"].value #
Cont_sweep = sheet["B10"].value #


FSV.write(f"*RST")
FSV.write("SYST:DISP:UPD ON")
FSV.write("ADEM ON")
FSV.write(f"FREQ:CENT {Centre_frequency}MHz")
FSV.write(f"DISP:TRAC:Y:PDIV {Dev_PerDivision}kHz")
FSV.write(f"BAND:DEM {Demod_BW}kHz")
FSV.write(f"ADEM:AF:COUP {AF_Couple}")#Set AF coupling to AC, the frequency offset is automatically corrected.
# i.e. the trace is always symmetric with respect to the zero line
FSV.write(f"DISP:TRAC:Y:RLEV:OFFS {RefLev_offset}")
FSV.write(f"DISP:TRAC:Y:RLEV {RF_level}")
FSV.write(f"INP:ATT {Attenuation}")
FSV.write(f"{Trace_Peak}")
FSV.write(f"ADEM:MTIM {Demod_MT}ms")
FSV.write(f"INIT:CONT {Cont_sweep}")
FSV.query("*OPC?")
time.sleep(5)# wait 5 seconds for trace to stable
FSV.write(f"INIT:CONT OFF")
Deviation = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0
print(f"Deviation is {Deviation}kHz")



#FSV.query("INIT:*WAI")




FSV.close()


#FSV.write(f"DISP:TRAC:MODE MAXH")
