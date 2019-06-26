import sys
import os
import math
import visa
from openpyxl import load_workbook
import time
import serial
import re
import config
import datetime
from decimal import *
#print(decimal.__file__)


# below codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder
sys.path.insert(0, r'C:\Users\afan\kallithea\aaron test\midlevel')
sys.path.insert(0, r'C:\Users\afan\kallithea\aaron test\bottomlevel')
sys.path.insert(0, r'C:\Users\afan\kallithea\aaron test\toplevel')
# above codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder


class SpecAn(object):

    def __init__(self):

        self.rm = visa.ResourceManager()
        self.SP = self.rm.open_resource('TCPIP0::192.168.10.9::hislip0::INSTR') # Spec An
        self.book = load_workbook(filename = "Test_Setup.xlsx") # load an existing .xlsx file

    def FEP_Setup(self, freq):

        self.sheet = self.book["Freq_Error_Power"] # load excel sheet

        # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
        self.sheet.cell(row = 1, column = 2, value = freq) # write test frequency in this self.sheet
        self.book.save("Test_Setup.xlsx") # save existing .xlsx file
        # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

        self.Centre_frequency = self.sheet["B1"].value #
        self.Span_frequency = self.sheet["B2"].value #
        self.RBW = self.sheet["B3"].value #
        self.VBW = self.sheet["B4"].value #
        self.RF_level = self.sheet["B5"].value #
        self.Attenuation = self.sheet["B6"].value #
        self.RefLev_offset = self.sheet["B7"].value#
        self.Trace_peak = self.sheet["B8"].value #
        self.Transducer1 = self.sheet["B9"].value # cable
        self.Trans1_ON = self.sheet["B10"].value
        self.Transducer2 = self.sheet["B11"].value # 30dB attenuator
        self.Trans2_ON = self.sheet["B12"].value
        self.Transducer3 = self.sheet["B13"].value # High pass filter
        self.Trans3_ON = self.sheet["B14"].value
        self.Limit_line_1 = self.sheet["B15"].value # ASNZS4365:2011
        self.Limit_line_1_ON = self.sheet["B16"].value # ASNZS4365:2011
        self.Limit_line_2 = self.sheet["B17"].value # ASNZS4295:2015
        self.Limit_line_2_ON = self.sheet["B18"].value # ASNZS4295:2015
        self.Sweep_points= self.sheet["B19"].value # get sweep points

        self.SP.write(f"*RST")
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write(f"FREQ:CENT {self.Centre_frequency}MHz")
        self.SP.write(f"FREQ:SPAN {self.Span_frequency}Hz")
        self.SP.write(f"BAND {self.RBW}Hz")
        self.SP.write(f"BAND:VID {self.VBW}Hz")
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {self.RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {self.RF_level}")
        self.SP.write(f"INP:ATT {self.Attenuation}")
        self.SP.write(f"{self.RefLev_offset}")
        self.SP.write(f"{self.Trace_peak}")
        self.SP.write(f"CORR:TRAN:SEL '{self.Transducer1}'")
        self.SP.write(f"CORR:TRAN {self.Trans1_ON} ")
        self.SP.write(f"CORR:TRAN:SEL '{self.Transducer2}'")
        self.SP.write(f"CORR:TRAN {self.Trans2_ON}")
        self.SP.write(f"CORR:TRAN:SEL '{self.Transducer3}'")
        self.SP.write(f"CORR:TRAN {self.Trans3_ON}")
        self.SP.write(f"CALC:LIM:NAME '{self.Limit_line_1}'")
        self.SP.write(f"CALC:LIM:UPP:STAT {self.Limit_line_1_ON}")
        self.SP.write(f"CALC:LIM:NAME '{self.Limit_line_2}'")
        self.SP.write(f"CALC:LIM:UPP:STAT {self.Limit_line_2_ON}")
        self.SP.write(f"SWE:POIN {self.Sweep_points}")


    def DeMod_Setup(self, freq):
        self.sheet = self.book["Analog_Demod"]
        # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
        self.sheet.cell(row = 1, column = 2, value = freq) # write test frequency in this self.sheet
        self.book.save("Test_Setup.xlsx") # save existing .xlsx file
        # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

        # following code is to initialize FSV
        self.Centre_frequency = self.sheet["B1"].value #
        self.Dev_PerDivision = self.sheet["B2"].value #
        self.Demod_BW = self.sheet["B3"].value #
        self.AF_Couple = self.sheet["B4"].value #
        self.RF_level = self.sheet["B5"].value #
        self.Attenuation = self.sheet["B6"].value # get attenuation
        self.RefLev_offset = self.sheet["B7"].value# get RFlevel offset
        self.Trace_Peak = self.sheet["B8"].value # get trace to Pos peak
        self.Demod_MT = self.sheet["B9"].value #
        self.Cont_sweep = self.sheet["B10"].value #
        self.SP.write(f"*RST")
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write("ADEM ON")
        self.SP.write(f"FREQ:CENT {self.Centre_frequency}MHz")
        self.SP.write(f"DISP:TRAC:Y:PDIV {self.Dev_PerDivision}kHz")
        self.SP.write(f"BAND:DEM {self.Demod_BW}kHz")
        self.SP.write(f"ADEM:AF:COUP {self.AF_Couple}")#Set AF coupling to AC, the frequency offset is automatically corrected.
        # i.e. the trace is always symmetric with respect to the zero line
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {self.RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {self.RF_level}")
        self.SP.write(f"INP:ATT {self.Attenuation}")
        self.SP.write(f"{self.Trace_Peak}")
        self.SP.write(f"ADEM:MTIM {self.Demod_MT}ms")
        self.SP.write(f"INIT:CONT {self.Cont_sweep}")
        # above code is to initialize FSV

    def ACP_Setup(self, freq):
        self.sheet = self.book["ACP"]
        # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
        self.sheet.cell(row = 1, column = 2, value = freq) # write test frequency in this self.sheet
        self.book.save("Test_Setup.xlsx") # save existing .xlsx file
        # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

        # following code is to initialize FSV to complete the test
        self.Centre_frequency = self.sheet["B1"].value # get all parameters from Test_Setup.xlsx
        self.Span_frequency = self.sheet["B2"].value #
        self.RBW = self.sheet["B3"].value #
        self.VBW = self.sheet["B4"].value #
        self.RF_level = self.sheet["B5"].value #
        self.Attenuation = self.sheet["B6"].value #
        self.RefLev_offset = self.sheet["B7"].value#
        self.Trace_RMS = self.sheet["B8"].value #

        self.Tx_CHBW = self.sheet["B10"].value
        self.AJ_CHBW = self.sheet["B11"].value
        self.AT_CHBW = self.sheet["B12"].value
        self.AJ_CHNUM = self.sheet["B13"].value
        self.AJ_SPACE = self.sheet["B14"].value
        self.AT_SPACE = self.sheet["B15"].value
        self.Power_Mode = self.sheet["B16"].value
        self.Ave_number = self.sheet["B17"].value

        self.SP.write(f"*RST")
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write("CALC:MARK:FUNC:POW:SEL ACP")

        self.SP.write(f"FREQ:CENT {self.Centre_frequency}MHz") # set all parameters
        self.SP.write(f"FREQ:SPAN {self.Span_frequency}kHz")
        self.SP.write(f"BAND {self.RBW}Hz")
        self.SP.write(f"BAND:VID {self.VBW}Hz")
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {self.RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {self.RF_level}")
        self.SP.write(f"INP:ATT {self.Attenuation}")
        self.SP.write(f"{self.Trace_RMS}")

        self.SP.write(f"POW:ACH:BWID:CHAN1 {self.Tx_CHBW}kHz")
        self.SP.write(f"POW:ACH:BWID:ACH {self.AJ_CHBW}kHz")
        self.SP.write(f"POW:ACH:BWID:ALT1 {self.AT_CHBW}kHz")
        self.SP.write(f"POW:ACH:ACP {self.AJ_CHNUM}")
        self.SP.write(f"POW:ACH:SPAC {self.AJ_SPACE}kHz")
        self.SP.write(f"POW:ACH:SPAC:ALT1 {self.AT_SPACE}kHz")
        self.SP.write(f"POW:ACH:MODE {self.Power_Mode}")
        self.SP.write(f"SWE:COUN {self.Ave_number}")
        self.SP.write(f"CALC:MARK:FUNC:POW:MODE WRIT")
        self.SP.write(f"DISP:TRAC:MODE AVER")


    def CSE_Setup(self, sub_range, limit_line):

        if sub_range == 1:
            self.sheet = self.book["Cond_Spurious_1"]
        elif sub_range == 2:
            self.sheet = self.book["Cond_Spurious_2"]
        elif sub_range == 3:
            self.sheet = self.book["Cond_Spurious_3"]
        elif sub_range == 4:
            self.sheet = self.book["Cond_Spurious_4"]
        elif sub_range == 5:
                self.sheet = self.book["Cond_Spurious_5"]
        else:
            raise ValueError('sub_range only accept integer 1 or 5')


        if limit_line == 2:
            self.sheet.cell(row = 16, column = 2, value = "ON") # write test frequency in this sheet
            self.sheet.cell(row = 18, column = 2, value = "OFF") # write test frequency in this sheet
            self.book.save("Test_Setup.xlsx")
        elif limit_line == 3:
            self.sheet.cell(row = 16, column = 2, value = "OFF") # write test frequency in this sheet
            self.sheet.cell(row = 18, column = 2, value = "ON") # write test frequency in this sheet
            self.book.save("Test_Setup.xlsx")
        else:
            raise ValueError('limit_line only accept integer 2 or 3')

        self.start_frequency = self.sheet["B1"].value # get start frequency value from sheet
        self.stop_frequency = self.sheet["B2"].value # get stop frequency value from sheet
        self.RBW = self.sheet["B3"].value # get RBW value from sheet
        self.VBW = self.sheet["B4"].value # get VBW value from sheet
        self.RF_level = self.sheet["B5"].value # get RF_level
        self.Attenuation = self.sheet["B6"].value # get attenuation
        self.RefLev_offset = self.sheet["B7"].value# get RFlevel offset
        self.Trace_peak = self.sheet["B8"].value # get trace to Pos peak
        self.Transducer1 = self.sheet["B9"].value # cable
        self.Trans1_ON = self.sheet["B10"].value
        self.Transducer2 = self.sheet["B11"].value # 30dB attenuator
        self.Trans2_ON = self.sheet["B12"].value
        self.Transducer3 = self.sheet["B13"].value # High pass filter
        self.Trans3_ON = self.sheet["B14"].value
        self.Limit_line_1 = self.sheet["B15"].value # ASNZS4365:2011
        self.Limit_line_1_ON = self.sheet["B16"].value # ASNZS4365:2011
        self.Limit_line_2 = self.sheet["B17"].value # ASNZS4365:2011
        self.Limit_line_2_ON = self.sheet["B18"].value # ASNZS4365:2011
        self.Sweep_points= self.sheet["B19"].value # get sweep points


        self.SP.write(f"*RST")
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write(f"FREQ:STAR {self.start_frequency}MHz")
        self.SP.write(f"FREQ:STOP {self.stop_frequency}MHz")
        self.SP.write(f"BAND {self.RBW}kHz")
        self.SP.write(f"BAND:VID {self.VBW}kHz")
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {self.RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {self.RF_level}")
        self.SP.write(f"INP:ATT {self.Attenuation}")
        self.SP.write(f"{self.RefLev_offset}")
        self.SP.write(f"{self.Trace_peak}")
        self.SP.write(f"CORR:TRAN:SEL '{self.Transducer1}'")
        self.SP.write(f"CORR:TRAN {self.Trans1_ON} ")
        self.SP.write(f"CORR:TRAN:SEL '{self.Transducer2}'")
        self.SP.write(f"CORR:TRAN {self.Trans2_ON}")
        self.SP.write(f"CORR:TRAN:SEL '{self.Transducer3}'")
        self.SP.write(f"CORR:TRAN {self.Trans3_ON}")
        self.SP.write(f"CALC:LIM:NAME '{self.Limit_line_1}'")
        self.SP.write(f"CALC:LIM:UPP:STAT {self.Limit_line_1_ON}")
        self.SP.write(f"CALC:LIM:NAME '{self.Limit_line_2}'")
        self.SP.write(f"CALC:LIM:UPP:STAT {self.Limit_line_2_ON}")
        self.SP.write(f"SWE:POIN {self.Sweep_points}")
        self.SP.write(f"DISP:TRAC:MODE MAXH")

    def TranP_Setup(self, freq):
        self.sheet = self.book["Tran_Perform"]
        # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
        self.sheet.cell(row = 1, column = 2, value = freq) # write test frequency in this self.sheet
        self.book.save("Test_Setup.xlsx") # save existing .xlsx file
        # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

        # following code is to initialize FSV
        self.Centre_frequency = self.sheet["B1"].value #
        self.Dev_PerDivision = self.sheet["B2"].value #
        self.Demod_BW = self.sheet["B3"].value #
        self.AF_Couple = self.sheet["B4"].value #
        self.RF_level = self.sheet["B5"].value #
        self.Attenuation = self.sheet["B6"].value # get attenuation
        self.RefLev_offset = self.sheet["B7"].value# get RFlevel offset
        self.Trace_Peak = self.sheet["B8"].value # get trace to Pos peak
        self.Demod_MT = self.sheet["B9"].value #
        self.Cont_sweep = self.sheet["B10"].value #
        self.Trigger_offset = self.sheet["B11"].value

        self.SP.write(f"*RST")
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write("ADEM ON")
        self.SP.write(f"FREQ:CENT {self.Centre_frequency}MHz")
        self.SP.write(f"DISP:TRAC:Y:PDIV {self.Dev_PerDivision}kHz")
        self.SP.write(f"BAND:DEM {self.Demod_BW}kHz")
        self.SP.write(f"ADEM:AF:COUP {self.AF_Couple}")#Set AF coupling to AC, the frequency offset is automatically corrected.
        # i.e. the trace is always symmetric with respect to the zero line
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {self.RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {self.RF_level}")
        self.SP.write(f"INP:ATT {self.Attenuation}")
        self.SP.write(f"{self.Trace_Peak}")
        self.SP.write(f"ADEM:MTIM {self.Demod_MT}ms")
        self.SP.write(f"INIT:CONT {self.Cont_sweep}")
        self.SP.write(f"TRIG:HOLD {self.Trigger_offset}")

        # above code is to initialize FSV




    def get_centfreq(self):
        return self.Centre_frequency # useful for ACP test


    def write(self, str):
        self.SP.write(str)

    def query(self, str):
        return self.SP.query(str)


    def get_FEP_result(self):
        self.SP.write("CALC:MARK1:MAX")
        self.Frequency = self.SP.query("CALC:MARK1:X?")
        self.Level = float(self.SP.query("CALC:MARK1:Y?"))
        self.Frequency_error = float(self.Frequency) - float(self.Centre_frequency)*1e6
        self.indication = (self.SP.query("*OPC?")).replace("1","Completed.")

        return {'F':self.Frequency_error, 'P':self.Level, 'I':self.indication}

    def get_CSE_result(self):
        self.SP.write("CALC:MARK1:MAX")
        self.Frequency = float(self.SP.query("CALC:MARK1:X?"))/1e6
        self.Level = float(self.SP.query("CALC:MARK1:Y?"))
        self.indication = (self.SP.query("*OPC?")).replace("1","Completed.")
        return {'F':self.Frequency, 'P':self.Level, 'I':self.indication}

    def screenshot(self, file_name):
        self.SP.write("HCOP:DEV:LANG PNG")
        self.SP.write("HCOP:CMAP:DEF4")
        self.SP.write(f"MMEM:NAME \'c:\\temp\\Dev_Screenshot.png\'")
        self.SP.write("HCOP:IMM")
        self.SP.query("*OPC?")

        file_data = self.SP.query_binary_values(f"MMEM:DATA? \'c:\\temp\\Dev_Screenshot.png\'", datatype='s',)[0]
        new_file = open(f"c:\\Temp\\{file_name}.png", "wb")# extract file_name string using (f"{}")
        new_file.write(file_data)
        new_file.close()
        print(f"saved to PC c:\\Temp\\{file_name}.png\n") # extract file_name string using (f"{}")

    def close(self):
        self.SP.close()

class SigGen(object):

    def __init__(self, address):
        self.rm = visa.ResourceManager()
        self.SG = self.rm.open_resource(address)
        self.book = load_workbook(filename = "Test_Setup.xlsx") # load an existing .xlsx file


    def MAD_Setup(self):
        self.sheet = self.book["Max_Deviation"]
        # following code is to initialize SML
        self.Frequency_AF = self.sheet["G1"].value #
        self.Level_AF = self.sheet["G2"].value #
        self.AF_output_on = self.sheet["G3"].value #

        self.SG.write(f"*RST")
        self.SG.write("SYST:DISP:UPD ON")
        self.SG.write(f":FM:INT:FREQ {self.Frequency_AF}kHz")
        self.SG.write(f":OUTP2:VOLT {self.Level_AF}mV")
        self.SG.write(f":OUTP2 {self.AF_output_on}")

    def TranP_Setup(self, freq):
        self.sheet = self.book["Tran_Perform"]
        # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
        self.sheet.cell(row = 1, column = 11, value = freq) # write test frequency in this self.sheet
        self.book.save("Test_Setup.xlsx") # save existing .xlsx file
        # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

        # following code is to initialize SML
        self.Frequency_RF = self.sheet["G1"].value #
        self.Level_RF = self.sheet["G2"].value #
        self.Frequency_AF = self.sheet["G3"].value #
        self.Deviation = self.sheet["G4"].value #
        self.Mod_state = self.sheet["G5"].value #
        self.RF_power_on = self.sheet["G6"].value #

        self.SG.write(f"*RST")
        self.SG.write("SYST:DISP:UPD ON")
        self.SG.write(f"FREQ {self.Frequency_RF}MHz")
        self.SG.write(f":POW:UNIT dBm")
        self.SG.write(f":POW {self.Level_RF}dBm")
        self.SG.write(f":FM:INT:FREQ {self.Frequency_AF}kHz")
        self.SG.write(f":FM:DEV {self.Deviation}kHz")
        self.SG.write(f":FM:STAT {self.Mod_state}")
        self.SG.write(f":OUTP1 {self.RF_power_on}")

    def Unwanted_Signal(self, freq):
        self.sheet = self.book["Rx_Siggen_Setting"]
        # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
        self.sheet.cell(row = 8, column = 3, value = freq) # write test frequency in this self.sheet
        self.book.save("Test_Setup.xlsx") # save existing .xlsx file
        # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

        self.Frequency_RF = self.sheet["C8"].value #
        self.Level_RF = self.sheet["C9"].value #
        self.Frequency_AF = self.sheet["C10"].value #
        self.Deviation = self.sheet["C11"].value #
        self.Mod_state = self.sheet["C12"].value #
        self.RF_power_on = self.sheet["C13"].value #

        self.SG.write(f"*RST")
        self.SG.write("SYST:DISP:UPD ON")
        self.SG.write(f"FREQ {self.Frequency_RF}MHz")
        self.SG.write(f":POW:UNIT dBuV")
        self.SG.write(f":POW {self.Level_RF}dBuV")
        self.SG.write(f":FM:INT:FREQ {self.Frequency_AF}kHz")
        self.SG.write(f":FM:DEV {self.Deviation}kHz")
        self.SG.write(f":FM:STAT {self.Mod_state}")
        self.SG.write(f":OUTP1 {self.RF_power_on}")

    def Wanted_Signal(self, freq):
        self.sheet = self.book["Rx_Siggen_Setting"]
        # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
        self.sheet.cell(row = 1, column = 3, value = freq) # write test frequency in this self.sheet
        self.book.save("Test_Setup.xlsx") # save existing .xlsx file
        # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

        self.Frequency_RF = self.sheet["C1"].value #
        self.Level_RF = self.sheet["C2"].value #
        self.Frequency_AF = self.sheet["C3"].value #
        self.Deviation = self.sheet["C4"].value #
        self.Mod_state = self.sheet["C5"].value #
        self.RF_power_on = self.sheet["C6"].value #

        self.SG.write(f"*RST")
        self.SG.write("SYST:DISP:UPD ON")
        self.SG.write(f"FREQ {self.Frequency_RF}MHz")
        self.SG.write(f":POW:UNIT dBuV")
        self.SG.write(f":POW {self.Level_RF}dBuV")
        self.SG.write(f":FM:INT:FREQ {self.Frequency_AF}kHz")
        self.SG.write(f":FM:DEV {self.Deviation}kHz")
        self.SG.write(f":FM:STAT {self.Mod_state}")
        self.SG.write(f":OUTP1 {self.RF_power_on}")


    def Set_Timeout(self, ms):
        self.SG.timeout = ms # useful on CMS for Rx test, pyvisa parameter


    def Lev_AF(self):
        return self.Level_AF # useful for Max_deviation test

    def Lev_RF(self):
        return self.Level_RF # useful for Rx test

    def write(self, str):
        self.SG.write(str)

    def query(self, str):
        return self.SG.query(str)

    def close(self):
        self.SG.close()


class Radio(object):
    def __init__(self, com):
        self.ESC_CHAR = 0x7D # Escape char
        self.STX_CHAR =  0x7E # Start of packet
        self.ETX_CHAR = 0x7F # End of packet
        self.CHECKSUM_XOR_MASK = 0xFF

        self.payload0 = bytearray(b'\xB4\x88\x2a\x80')# set carrie freq 136MHz basd on 12.5kHz calculation
        self.payload1 = bytearray(b'\xB4\x93\x03')# set 25W power
        self.payload2 = bytearray(b'\xA6\x01') # PTT ON
        self.payload3 = bytearray(b'\xB4\x91\x03\xE8')# 1kHz audio on
        self.payload4 = bytearray(b'\xB4\x91\x0B\xB8')# 3kHz audio on
        self.payload5 = bytearray(b'\xB4\x91\x00\x00')# audio off
        self.payload6 = bytearray(b'\xB4\x82\x10')# selcall tone on
        self.payload7 = bytearray(b'\xB4\x82\x0f')# selcall tone off
        self.payload8 = bytearray(b'\xA6\x00') # PTT off

        self.port = com
        self.baudrate = 115200
        self.timeout = None
        self.ser = serial.Serial(port=self.port, baudrate=self.baudrate, timeout=self.timeout)  # open serial port


    def Packet_Gen(self, payload):
        # Header
        tx_array = bytearray((self.STX_CHAR, 2, len(payload)))

        # Escape the payload, append it to the output buffer, make the checksum
        chksum = 0
        for i, b in enumerate(payload):
            # The first word of the payload is always the command ID.
            # The MSB of the command ID is always 1. The radio just
            # clears it in the ACK, without doing anything else with it.
            if i == 0:
                bb = (0x80 | b) & 0xFF
            else:
                bb = b & 0xFF
            chksum ^= bb
            if bb in (self.ESC_CHAR, self.STX_CHAR, self.ETX_CHAR):
                tx_array.append(self.ESC_CHAR)
                tx_array.append(bb ^ 0x20)
            else:
                tx_array.append(bb)

        # Trailer
        tx_array.append(chksum ^ self.CHECKSUM_XOR_MASK)
        tx_array.append(self.ETX_CHAR)

        return tx_array

    def Set_Freq(self, freq):
        a = ((hex(int((freq*Decimal(1e6))/(Decimal(12.5*1e3))))[2]+hex(int((freq*Decimal(1e6))/Decimal((12.5*1e3))))[3]))# calculate first HEX of input frequency
        b = ((hex(int((freq*Decimal(1e6))/(Decimal(12.5*1e3))))[4]+hex(int((freq*Decimal(1e6))/Decimal((12.5*1e3))))[5]))# calculate second HEX of input frequency
        self.payload0[2] = int(a,16) # assign first DEC (transferred from HEX) to the third number of paylaod0
        self.payload0[3] = int(b,16) # assign second DEC (transferred from HEX) to the fourth number of paylaod0
        self.ser.write(self.Packet_Gen(self.payload0))
        print(f"radio frequency has been set to {freq} MHz.")

    def Set_Pow(self, pow):
        if pow == "low":
            self.payload1[2] = 0
        elif pow == "high":
            self.payload1[2] = 3
        else:
            print("Set_Pow() only accept 'low' or 'high'")
        print(f"radio power has been set to {pow}.")

        self.ser.write(self.Packet_Gen(self.payload1))

    def Radio_On(self):
        self.ser.write(self.Packet_Gen(self.payload2))
        print("Radio's on")

    def Radio_Off(self):
        self.ser.write(self.Packet_Gen(self.payload8))
        print("Radio's off")



    # def Radio_Control(self):
    #     # time.sleep(0.5)
    #     self.ser.write(self.Packet_Gen(payload))
    #     print("one command sent.") # set a indicator showing one command has been sent to radio

    def Radio_close(self):
        self.ser.close()
        print("Serial session is closed.")

class Excel(object):
    def __init__(self, file_name):
        self.file = load_workbook(filename = file_name) # load Test_Result.xlsx

    def get_sheet(self, sheet_name):
        self.sheet = self.file[sheet_name]

    def write(self, row, column, value):
        self.sheet.cell(row = row, column = column, value = value)

    def save(self, file_name):
        self.file.save(file_name)

def Tx_Frequency_error_Carrier_power():

    Result_sheet.get_sheet("Ferror_Pow")
    FSV.FEP_Setup(470)
    FSV.query('*OPC?')


    CP50.Set_Freq(470)
    CP50.Set_Pow("high")
    CP50.Radio_On()
    time.sleep(2)
    FSV.write("DISP:TRAC:MODE MAXH")
    time.sleep(3)
    FSV.write("DISP:TRAC:MODE VIEW")
    CP50.Radio_Off()
    dict = FSV.get_FEP_result()
    print(f"Frequency error and Carrier power test {dict['I']}")
    print(f"Frequency error:{dict['F']}Hz")
    print(f"Carrier power:{dict['P']}dBm")

    Result_sheet.write(row = 2, column = 1, value = dict['F'])
    Result_sheet.write(row = 2, column = 2, value = dict['P'])
    Result_sheet.save("Test_Result.xlsx")

    # FSV.close()
    # CP50.Radio_close()


def Tx_Max_deviation():
    AF_list = [100, 300, 500, 700, 900, 1000, 1300, 1500, 1700, 1900, 2000, 2300, 2550] # audio frequency list
    Reading_list = [] # empty list for deviation result storage
    # RSheet = Get_result_sheet("Test_Result.xlsx", "Max_Dev")
    Result_sheet.get_sheet("Max_Dev")
    SML.MAD_Setup()
    SML.query('*OPC?')
    FSV.DeMod_Setup(470)
    FSV.query('*OPC?')

    CP50.Set_Freq(470)
    CP50.Set_Pow("high")
    CP50.Radio_On()
    time.sleep(2)

    FSV.write(f"INIT:CONT OFF")

    Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0 #get the initial deviation value


    Level_AF = SML.Lev_AF()# initial Level_AF

    if Dev_Reading < 1.5:
        while Dev_Reading < 1.47:
            Level_AF = Level_AF+1
            SML.write(f":OUTP2:VOLT {Level_AF}mV")
            FSV.write(f"INIT:CONT ON")
            time.sleep(1)
            FSV.write(f"INIT:CONT OFF")
            time.sleep(1)
            Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0
    else:
        while Dev_Reading > 1.53:
            Level_AF = Level_AF-1
            SML.write(f":OUTP2:VOLT {Level_AF}mV")
            FSV.write(f"INIT:CONT ON")
            time.sleep(1)
            FSV.write(f"INIT:CONT OFF")
            time.sleep(1)
            Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0
    # above code is to find Audio output level satisfying standard condtion (around 1.5kHz deviation)

    FSV.write(f"INIT:CONT ON")
    SML.query('*OPC?')


    #following code is to vary audio frequency to complete the test
    Level_AF = 100*Level_AF # bring audio level up 20dB in one step according to standard
    SML.write(f":OUTP2:VOLT {Level_AF}mV")

    for i in range(0,13):
        print(f"At Audio frequency:{AF_list[i]}")
        SML.write(f":FM:INT:FREQ {AF_list[i]}Hz")
        SML.query('*OPC?')
        FSV.write(f"INIT:CONT OFF")
        time.sleep(1)
        Dev = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000
        Reading_list.append(Dev)
        FSV.query("*OPC?")
        FSV.write(f"INIT:CONT ON")
        print(f"Deviation is {Reading_list[i]}kHz")
        Result_sheet.write(row = i+2, column = 1, value = AF_list[i])
        Result_sheet.write(row = i+2, column = 2, value = Dev)


    Result_sheet.save("Test_Result.xlsx") # save existing .xlsx file
    CP50.Radio_Off()
    SML.write(":OUTP2 OFF")# turn off audio output at the end of the test
    indication = (FSV.query("*OPC?")).replace("1","Completed.")
    print(f"Maximum Deviation test {indication}")

    # FSV.close()
    # SML.close()
    # CP50.Radio_close()

def Tx_Adjacent_channel_power():
    Result_sheet.get_sheet("ACP")
    SML.Tx_Setup()
    SML.query('*OPC?')
    FSV.DeMod_Setup(470)
    FSV.query('*OPC?')

    CP50.Set_Freq(470)
    CP50.Set_Pow("high")
    CP50.Radio_On()
    time.sleep(2)

    FSV.write(f"INIT:CONT OFF")

    Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0 #get the initial deviation value


    Level_AF = SML.Lev_AF()# initial Level_AF

    if Dev_Reading < 1.5:
        while Dev_Reading < 1.47:
            Level_AF = Level_AF+1
            SML.write(f":OUTP2:VOLT {Level_AF}mV")
            FSV.write(f"INIT:CONT ON")
            time.sleep(1)
            FSV.write(f"INIT:CONT OFF")
            time.sleep(1)
            Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0
    else:
        while Dev_Reading > 1.53:
            Level_AF = Level_AF-1
            SML.write(f":OUTP2:VOLT {Level_AF}mV")
            FSV.write(f"INIT:CONT ON")
            time.sleep(1)
            FSV.write(f"INIT:CONT OFF")
            time.sleep(1)
            Dev_Reading = float(FSV.query("CALC:MARK:FUNC:ADEM:FM? MIDD"))/1000.0
    # above code is to find Audio output level satisfying standard condtion (around 1.5kHz deviation)

    FSV.write(f"INIT:CONT ON")
    FSV.query('*OPC?')

    FSV.ACP_Setup(470)
    time.sleep(8)
    FSV.write(f"DISP:TRAC:MODE VIEW")


    ACP = FSV.query("CALC:MARK:FUNC:POW:RES? ACP")
    print(ACP)
    LIST = re.findall(r'\d+\.\d+', ACP)
    Result_sheet.write(row = 2, column = 1, value = FSV.get_centfreq())
    Result_sheet.write(row = 2, column = 2, value = float(LIST[0]))
    Result_sheet.write(row = 2, column = 3, value = -float(LIST[1]))
    Result_sheet.write(row = 2, column = 4, value = -float(LIST[2]))

    Result_sheet.save("Test_Result.xlsx") # save existing .xlsx file
    CP50.Radio_Off()
    SML.write(":OUTP2 OFF")# turn off audio output at the end of the test
    indication = (FSV.query("*OPC?")).replace("1","Completed.")
    print(f"ACP test {indication}")

    # FSV.close()
    # SML.close()
    # CP50.Radio_close()

def Tx_Conducted_spurious_emissions():
    while True:
        print("1. low subranges test without high pass filter, make sure no filter installed:")
        print("2. high subranges test with high pass filter, make sure filter installed:")
        print("3. go back:")
        choice = input("> ")
        if choice == "1":
            for i in range(0, 3):# first 3 subranges without high pass filter
                FSV.CSE_Setup(sub_range=i+1, limit_line=config.limit_line_factor)
                FSV.query('*OPC?')
                CSE_operation(row=i+2, sub_range=i+1)
                FSV.screenshot('CSE0'+str(i+1))

        elif choice == "2":
            for i in range(3, 5):# second 2 subranges with high pass filter
                FSV.CSE_Setup(sub_range=i+1, limit_line=config.limit_line_factor)
                FSV.query('*OPC?')
                CSE_operation(row=i+2, sub_range=i+1)
                FSV.screenshot('CSE0'+str(i+1))

        elif choice == "3":
            break

        else:
            print("choice only take number 1, 2 or 3, try agin...")


    #FSV.close()
    #CP50.Radio_close()



def CSE_operation(row, sub_range):

    Result_sheet.get_sheet("Cond_Spur")
    CP50.Set_Freq(470)
    CP50.Set_Pow("high")
    CP50.Radio_On()
    time.sleep(3)
    FSV.write("DISP:TRAC:MODE VIEW")
    CP50.Radio_Off()
    dict = FSV.get_CSE_result()
    print(f"Conducted spurious emission test {dict['I']}")
    print(f"Marker frequency:{dict['F']}MHz")
    print(f"Marker power:{dict['P']}dBm")

    Result_sheet.write(row = row, column = 1, value = sub_range)
    Result_sheet.write(row = row, column = 2, value = dict['F'])
    Result_sheet.write(row = row, column = 3, value = dict['P'])
    Result_sheet.write(row = row, column = 4, value = -30)
    Result_sheet.save("Test_Result.xlsx")


def Tx_Transient_performance():

    SML.TranP_Setup(459.075)
    SML.query('*OPC?')
    FSV.TranP_Setup(459.075)
    FSV.query('*OPC?')
    time.sleep(2)

    FSV.write("TRIG:SOUR RFP")
    time.sleep(2)
    FSV.write("INIT:CONT OFF")
    time.sleep(2)
    FSV.write("INIT:CONM")
    time.sleep(2)

    CP50.Set_Freq(459.075)
    CP50.Set_Pow("high")
    CP50.Radio_On()
    time.sleep(1)
    CP50.Radio_Off()
    FSV.screenshot('TranP_Tx_on')
    time.sleep(3)
    # FSV.write("INIT:CONT ON")
    # FSV.write("TRIG:HOLD -0.022")# this value is still problematic
    # time.sleep(3)
    # CP50.Radio_On()
    # time.sleep(0.1)
    # CP50.Radio_Off()
    # time.sleep(2)
    # FSV.screenshot('TranP_Tx_off')

    #FSV.write("TRIG:SOUR IMM") # set FSV back to freerun


def Rx_Adjacent_channel_selectivity(freq, delta):# delta set frequency offset
    Result_sheet.get_sheet("ACS")
    SML.Wanted_Signal(freq=freq)
    SML.query('*OPC?')
    SMB.Unwanted_Signal(freq=freq+Decimal(delta))
    print(f"Inference frequency has been set to {freq+Decimal(delta)}MHz")
    SMB.query('*OPC?')
    CP50.Set_Freq(freq=freq)
    CMS.Set_Timeout(ms=10000)

    # read initial SINAD
    SINAD_data_str = CMS.query("SINAD:R?")
    SINAD_data_num = re.findall(r'\d+\.\d+', SINAD_data_str)[0]
    print(SINAD_data_num)
    Level_RF = SMB.Lev_RF()

    for i in range(0,100):
        if float(SINAD_data_num) > 14.0:
            Result_sheet.write(row = i+2, column = 1, value = Level_RF)
            Result_sheet.write(row = i+2, column = 2, value = SINAD_data_num)
            Level_RF = Level_RF + 1
            SMB.write(f":POW {Level_RF}dBuV")
            SMB.query('*OPC?')
            SINAD_data_str = CMS.query("SINAD:R?")
            SINAD_data_num = re.findall(r'\d+\.\d+', SINAD_data_str)[0]
            if SINAD_data_num != '0':
                SINAD_data_num = re.findall(r'\d+\.\d+', SINAD_data_str)[0]
                print(SINAD_data_num)
            else:
                break
        else:
            break

    ACS = Level_RF-33.5
    Result_sheet.write(row = 2, column = 10, value = ACS)
    Result_sheet.save("Test_Result.xlsx")
    return ACS

def CHSW_ACS():
    Start_F = Decimal(487)/Decimal(1) # Decimal module make sure float numbers addition yeilds correct value
    CP50_Result.get_sheet("ACS_SN02")

    for i in range(30, 40):
        ACS_high = Rx_Adjacent_channel_selectivity(freq=Start_F, delta=0.0125)
        CP50_Result.write(row = i+2, column = 1, value = Start_F)
        CP50_Result.write(row = i+2, column = 2, value = Start_F+Decimal(0.0125))
        CP50_Result.write(row = i+2, column = 3, value = ACS_high)


        ACS_low = Rx_Adjacent_channel_selectivity(freq=Start_F, delta=-0.0125)
        CP50_Result.write(row = i+2, column = 4, value = Start_F-Decimal(0.0125))
        CP50_Result.write(row = i+2, column = 5, value = ACS_low)
        Timestamp ='{:%d-%b-%Y %H:%M:%S}'.format(datetime.datetime.now())
        CP50_Result.write(row = i+2, column = 6, value = Timestamp)

        CP50_Result.save("CP50_Result.xlsx")
        Start_F += Decimal(0.0125) # frequency in MHz



def Rx_Spurious_response_immunity(freq, delta):
    Result_sheet.get_sheet("Spur_Res")
    SML.Wanted_Signal(freq=freq)
    SML.query('*OPC?')
    SMB.Unwanted_Signal(freq=freq+Decimal(delta))
    print(f"Inference frequency has been set to {freq+Decimal(delta)}MHz")
    SMB.query('*OPC?')
    CP50.Set_Freq(freq=freq)
    CMS.Set_Timeout(ms=10000)

    # read initial SINAD
    SINAD_data_str = CMS.query("SINAD:R?")
    SINAD_data_num = re.findall(r'\d+\.\d+', SINAD_data_str)[0]
    print(SINAD_data_num)
    Level_RF = SMB.Lev_RF()

    # below code block to test Spurious_Response
    for i in range(0,100):
        if float(SINAD_data_num) > 14.0:
            Result_sheet.write(row = i+2, column = 1, value = Level_RF)
            Result_sheet.write(row = i+2, column = 2, value = SINAD_data_num)
            Level_RF = Level_RF + 1
            SMB.write(f":POW {Level_RF}dBuV")
            SMB.query('*OPC?')
            SINAD_data_str = CMS.query("SINAD:R?")
            SINAD_data_num = re.findall(r'\d', SINAD_data_str)[0]# handle return value of 0
            if SINAD_data_num != '0':
                SINAD_data_num = re.findall(r'\d+\.\d+', SINAD_data_str)[0]
                print(SINAD_data_num)
            else:
                break
        else:
            break

    Spur_Res = Level_RF-33.5
    Result_sheet.write(row = 2, column = 3, value = Spur_Res)
    Result_sheet.save("Test_Result.xlsx")
    return Spur_Res


def CHSW_SR():

    Start_F = Decimal(451)/Decimal(1) # Decimal module make sure float numbers addition yeilds correct value
    CP50_Result.get_sheet("day2")


    for i in range(0, 5):
        Spur_Res = Rx_Spurious_response_immunity(freq=Start_F, delta=Decimal(-2*38.85))
        CP50_Result.write(row = i+2, column = 1, value = Start_F)
        CP50_Result.write(row = i+2, column = 2, value = Start_F+Decimal(-2*38.85))
        CP50_Result.write(row = i+2, column = 3, value = Spur_Res)
        Timestamp ='{:%d-%b-%Y %H:%M:%S}'.format(datetime.datetime.now())
        CP50_Result.write(row = i+2, column = 4, value = Timestamp)
        CP50_Result.save("CP50_Result.xlsx")
        Start_F += Decimal(0.0125) # frequency in MHz





#FSV = SpecAn()
CP50 = Radio("com7")
SML = SigGen('ASRL4::INSTR')
SMB = SigGen('USB0::0x0AAD::0x0054::106409::INSTR')
CMS = SigGen('GPIB0::24::INSTR')
Result_sheet = Excel("Test_Result.xlsx")
CP50_Result = Excel("CP50_Result.xlsx")
getcontext().prec = 10 # set 10 decimal values precision
