import visa
import time
from openpyxl import load_workbook

rm = visa.ResourceManager()
FSV = rm.open_resource('TCPIP0::192.168.10.9::hislip0::INSTR') # Spec An
SML = rm.open_resource('ASRL4::INSTR') # wanted Signal
AF_list = [100, 300, 500, 700, 900, 1000, 1300, 1500, 1700, 1900, 2000, 2300, 2550]


SML.clear()  # Clear instrument io buffers and status
FSV.clear()

FSV_file_write = load_workbook(filename = "Test_Setup.xlsx") # create a workbook from existing .xlsx file
sheet = FSV_file_write["Max_Deviation"] # load setup sheet in .xlsx to sheet

# following code is to initialize SML
Frequency_AF = sheet["G1"].value #
Level_AF = sheet["G2"].value #
AF_output_on = sheet["G3"].value #
SML.write(f"*RST")
SML.write("SYST:DISP:UPD ON")
SML.write(f":FM:INT:FREQ {Frequency_AF}kHz")
SML.write(f":OUTP2:VOLT {Level_AF}mV")
SML.write(f":OUTP2 {AF_output_on}")
SML.query('*OPC?')
time.sleep(1)
# above code is to initialize SML

# following code is to initialize FSV
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
time.sleep(1)
# above code is to initialize FSV



# following code is to find Audio output level satisfying standard condtion (around 1.5kHz deviation)
FSV.write(f"INIT:CONT OFF")
FSV.query('*OPC?')
time.sleep(1)
Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0 #get the initial deviation value

if Dev_Reading < 1.5:
    while Dev_Reading < 1.42:
        Level_AF = Level_AF+1
        SML.write(f":OUTP2:VOLT {Level_AF}mV")
        FSV.write(f"INIT:CONT ON")
        time.sleep(1)
        FSV.write(f"INIT:CONT OFF")
        time.sleep(1)
        Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0
else:
    while Dev_Reading > 1.58:
        Level_AF = Level_AF-1
        SML.write(f":OUTP2:VOLT {Level_AF}mV")
        FSV.write(f"INIT:CONT ON")
        time.sleep(1)
        FSV.write(f"INIT:CONT OFF")
        time.sleep(1)
        Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0
# above code is to find Audio output level satisfying standard condtion (around 1.5kHz deviation)

SML.query('*OPC?')

#following code is to vary audio frequency to complete the test
Level_AF = 100*Level_AF # bring audio level up 20dB in one step
SML.write(f":OUTP2:VOLT {Level_AF}mV")
for i in range(0,13):
    print(f"At Audio frequency:{AF_list[i]}")
    SML.write(f":FM:INT:FREQ {AF_list[i]}Hz")
    SML.query('*OPC?')
    FSV.write(f"INIT:CONT OFF")
    time.sleep(1)
    Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0
    FSV.query("*OPC?")
    FSV.write(f"INIT:CONT ON")
    print(f"Deviation is {Dev_Reading}kHz")
#above code is to vary audio frequency to complete the test
FSV.close()
