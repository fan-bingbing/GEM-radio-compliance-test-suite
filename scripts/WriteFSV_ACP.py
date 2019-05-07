import visa
from openpyxl import load_workbook

rm = visa.ResourceManager()
FSV = rm.open_resource('TCPIP0::192.168.10.9::hislip0::INSTR') # Spec An

FSV_file_write = load_workbook(filename = "FSV_Setup.xlsx") # create a workbook from existing .xlsx file
sheet = FSV_file_write["ACP"] # load setup sheet in .xlsx to sheet

Centre_frequency = sheet["B1"].value # get start frequency value from sheet
Span_frequency = sheet["B2"].value # get stop frequency value from sheet
RBW = sheet["B3"].value # get RBW value from sheet
VBW = sheet["B4"].value # get VBW value from sheet
RF_level = sheet["B5"].value # get RF_level
Attenuation = sheet["B6"].value # get attenuation
RefLev_offset = sheet["B7"].value# get RFlevel offset
Trace_RMS = sheet["B8"].value # get trace to Pos peak

Tx_CHBW = sheet["B10"].value
AJ_CHBW = sheet["B11"].value
AT_CHBW = sheet["B12"].value
AJ_CHNUM = sheet["B13"].value
AJ_SPACE = sheet["B14"].value
AT_SPACE = sheet["B15"].value
Power_Mode = sheet["B16"].value
Ave_number = sheet["B17"].value

FSV.write(f"*RST")
FSV.write("SYST:DISP:UPD ON")
FSV.write("CALC:MARK:FUNC:POW:SEL ACP")

FSV.write(f"FREQ:CENT {Centre_frequency}MHz")
FSV.write(f"FREQ:SPAN {Span_frequency}kHz")
FSV.write(f"BAND {RBW}Hz")
FSV.write(f"BAND:VID {VBW}Hz")
FSV.write(f"DISP:TRAC:Y:RLEV:OFFS {RefLev_offset}")
FSV.write(f"DISP:TRAC:Y:RLEV {RF_level}")
FSV.write(f"INP:ATT {Attenuation}")
FSV.write(f"{Trace_RMS}")

FSV.write(f"POW:ACH:BWID:CHAN1 {Tx_CHBW}kHz")
FSV.write(f"POW:ACH:BWID:ACH {AJ_CHBW}kHz")
FSV.write(f"POW:ACH:BWID:ALT1 {AT_CHBW}kHz")
FSV.write(f"POW:ACH:ACP {AJ_CHNUM}")
FSV.write(f"POW:ACH:SPAC {AJ_SPACE}kHz")
FSV.write(f"POW:ACH:SPAC:ALT1 {AT_SPACE}kHz")
FSV.write(f"POW:ACH:MODE {Power_Mode}")
FSV.write(f"SWE:COUN {Ave_number}")
FSV.write(f"CALC:MARK:FUNC:POW:MODE WRIT")
FSV.write(f"DISP:TRAC:MODE AVER")

FSV.close()
