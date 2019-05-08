from openpyxl import load_workbook

def Tx_Conducted_spurious_emissions_setup(setup_sheet):
    FSV_file_write = load_workbook(filename = "FSV_Setup.xlsx") # create a workbook from existing .xlsx file
    sheet = FSV_file_write[setup_sheet] # load setup sheet in .xlsx to sheet
    start_frequency = sheet["B1"].value # get start frequency value from sheet
    stop_frequency = sheet["B2"].value # get stop frequency value from sheet
    RBW = sheet["B3"].value # get RBW value from sheet
    VBW = sheet["B4"].value # get VBW value from sheet
    RF_level = sheet["B5"].value # get RF_level
    Attenuation = sheet["B6"].value # get attenuation
    RefLev_offset = sheet["B7"].value# get RFlevel offset
    Trace_peak = sheet["B8"].value # get trace to Pos peak
    Transducer1 = sheet["B9"].value # cable
    Trans1_ON = sheet["B10"].value
    Transducer2 = sheet["B11"].value # 30dB attenuator
    Trans2_ON = sheet["B12"].value
    Transducer3 = sheet["B13"].value # High pass filter
    Trans3_ON = sheet["B14"].value
    Limit_line = sheet["B15"].value # ASNZS4365:2011
    Sweep_points= sheet["B16"].value # get sweep points


    FSV.write(f"*RST")
    FSV.write("SYST:DISP:UPD ON")
    FSV.write(f"FREQ:STAR {start_frequency}MHz")
    FSV.write(f"FREQ:STOP {stop_frequency}MHz")
    FSV.write(f"BAND {RBW}kHz")
    FSV.write(f"BAND:VID {VBW}kHz")
    FSV.write(f"DISP:TRAC:Y:RLEV:OFFS {RefLev_offset}")
    FSV.write(f"DISP:TRAC:Y:RLEV {RF_level}")
    FSV.write(f"INP:ATT {Attenuation}")
    FSV.write(f"{RefLev_offset}")
    FSV.write(f"{Trace_peak}")
    FSV.write(f"CORR:TRAN:SEL '{Transducer1}'")
    FSV.write(f"CORR:TRAN {Trans1_ON} ")
    FSV.write(f"CORR:TRAN:SEL '{Transducer2}'")
    FSV.write(f"CORR:TRAN {Trans2_ON}")
    FSV.write(f"CORR:TRAN:SEL '{Transducer3}'")
    FSV.write(f"CORR:TRAN {Trans3_ON}")
    FSV.write(f"CALC:LIM:NAME '{Limit_line}'")
    FSV.write(f"CALC:LIM:UPP:STAT ON")
    FSV.write(f"SWE:POIN 3000")
    FSV.write(f"DISP:TRAC:MODE MAXH")

def Tx_Conducted_spurious_emissions_test():
    print("blank")
