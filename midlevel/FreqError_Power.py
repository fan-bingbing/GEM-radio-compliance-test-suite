import visa
import time
from openpyxl import load_workbook



def FreqError_Power(Test_frequency):
    rm = visa.ResourceManager()
    FSV = rm.open_resource('TCPIP0::192.168.10.9::hislip0::INSTR') # Spec An
    # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
    FSV_file_write = load_workbook(filename = "Test_Setup.xlsx") # load an existing .xlsx file
    sheet = FSV_file_write["Freq_Error_Power"] # load existing sheet named "ACP"
    sheet.cell(row = 1, column = 2, value = Test_frequency) # write test frequency in this sheet
    FSV_file_write.save("Test_Setup.xlsx") # save existing .xlsx file
    # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

    FSV_file_write = load_workbook(filename = "Test_Setup.xlsx") # create a workbook from existing .xlsx file
    sheet = FSV_file_write["Freq_Error_Power"] # load setup sheet in .xlsx to sheet
    Centre_frequency = sheet["B1"].value #
    Span_frequency = sheet["B2"].value #
    RBW = sheet["B3"].value #
    VBW = sheet["B4"].value #
    RF_level = sheet["B5"].value #
    Attenuation = sheet["B6"].value #
    RefLev_offset = sheet["B7"].value#
    Trace_peak = sheet["B8"].value #
    Transducer1 = sheet["B9"].value # cable
    Trans1_ON = sheet["B10"].value
    Transducer2 = sheet["B11"].value # 30dB attenuator
    Trans2_ON = sheet["B12"].value
    Transducer3 = sheet["B13"].value # High pass filter
    Trans3_ON = sheet["B14"].value
    Limit_line_1 = sheet["B15"].value # ASNZS4365:2011
    Limit_line_1_ON = sheet["B16"].value # ASNZS4365:2011
    Limit_line_2 = sheet["B17"].value # ASNZS4295:2015
    Limit_line_2_ON = sheet["B18"].value # ASNZS4295:2015
    Sweep_points= sheet["B19"].value # get sweep points

    FSV.write(f"*RST")
    FSV.write("SYST:DISP:UPD ON")
    FSV.write(f"FREQ:CENT {Centre_frequency}MHz")
    FSV.write(f"FREQ:SPAN {Span_frequency}Hz")
    FSV.write(f"BAND {RBW}Hz")
    FSV.write(f"BAND:VID {VBW}Hz")
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
    FSV.write(f"CALC:LIM:NAME '{Limit_line_1}'")
    FSV.write(f"CALC:LIM:UPP:STAT {Limit_line_1_ON}")
    FSV.write(f"CALC:LIM:NAME '{Limit_line_2}'")
    FSV.write(f"CALC:LIM:UPP:STAT {Limit_line_2_ON}")
    FSV.write(f"SWE:POIN {Sweep_points}")
    FSV.write(f"DISP:TRAC:MODE MAXH")
    time.sleep(5)# wait 5 seconds for trace to stable
    FSV.write(f"DISP:TRAC:MODE VIEW")
    FSV.query("*OPC?")

    FSV.write("CALC:MARK1:MAX")
    Frequency = FSV.query("CALC:MARK1:X?")
    Level = FSV.query("CALC:MARK1:Y?")
    Frequency_error = float(Frequency) - float(Centre_frequency)*1e6
    indication = (FSV.query("*OPC?")).replace("1","Frequency error test Completed") # replace return character "1" to "completed"
    FSV.close()

    return {'Frequency_error':Frequency_error, 'Carrier_power':Level, 'Indication':indication}
    # return multiple values using dictionary
