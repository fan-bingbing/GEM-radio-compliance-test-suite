"""
GME radio compliance test suite scripts for class and function definition .
Copyright (C) 2019 Standard Communications Pty Ltd (GME). All rights reserved.
"""

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
from decimal import * # should aviod wildcard import mentioned in PEP08
import sounddevice as sd
import numpy as np
import matplotlib.pyplot as plt
from scipy.fftpack import fft



# below codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder
sys.path.insert(0, r'C:\Users\afan\kallithea\aaron test\midlevel')
sys.path.insert(0, r'C:\Users\afan\kallithea\aaron test\bottomlevel')
sys.path.insert(0, r'C:\Users\afan\kallithea\aaron test\toplevel')
# above codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder

class SoundCard(object):
    def __init__(self):
        self.gain = 10
        self.duration = 0.1 # sample record duration
        self.sample_rate = 44100
        self.device = ('Microphone (Sound Blaster Play!, MME',
                       'Speakers (Sound Blaster Play! 3, MME')
        sd.query_devices(self.device[0])
        sd.default.device = self.device
        sd.default.channels = 1
        sd.default.samplerate = self.sample_rate

    def get_sample(self):
        x1 = sd.rec(int(self.duration*self.sample_rate), dtype='float32', blocking=True) # record 0.1 seconds input signal from microphone
        x2 = self.gain*x1[int(self.duration*self.sample_rate/1000):int(self.duration*self.sample_rate)] # take effective samples
        x3 = x2[0:int(80*(self.sample_rate/1000))] # take first 80ms duration samples
        x4 = x3.flatten()# transfer 2D array to 1D, ready for fft operation, fft function can only take 1D array as input
        t = 80*np.linspace(0, 1, np.size(x4))# generate 80ms time axis VS samples for first 80ms

        X = fft(x4) # fft transform
        n = np.size(t)
        fr = (self.sample_rate/2)*np.linspace(0, 1, int(n/2)) # generate frequecny axis array
        X_m = (2/n)*abs(X[0:np.size(fr)])
        return {'Time_level':x4, 'Time':t, 'Freq_level':X_m, 'Freq':fr}

    def get_SINAD(self, dict):
        X_mmax = np.max(dict['Freq_level']) # find maximum value in numpy array
        fr_index = np.where(dict['Freq_level'] == X_mmax) # locate index in numpy array
        fr_max = dict['Freq'][fr_index] # find corresponding frequency of maximum value

        mask = np.full(np.size(dict['Freq']), 1) # initialize a mask numpy array with size=fr, value=1

        mask[fr_index[0]]=0 # make notch filter mask by assigning some points=0 where certain bandwidth of carrier was covered
        for i in range(1,41): # notch filter cover 40 bins
            mask[fr_index[0]-i]=0
            mask[fr_index[0]+i]=0

        SND = np.sum(dict['Freq_level']) # add all points value to get SND
        ND = np.sum(dict['Freq_level'] * mask) # apply notch filter, get ND
        SINAD = 20 * np.log10(SND/ND) # calculate SINAD
        return SINAD

    def live_plots(self):
        dict = self.get_sample()
        fig, (ax0, ax1) = plt.subplots(nrows=2)
        ax0.set_title('Sinusoidal Signal')
        ax0.set_xlabel('Time(ms)')
        ax0.set_ylabel('Amplitude')
        ax0.set_ylim(-2, 2)
        line0, = ax0.plot(dict['Time'], dict['Time_level']) # time VS level plot

        ax1.set_title('Magnitude Spectrum')
        ax1.set_xlabel('Frequency(Hz)')
        ax1.set_ylabel('Magnitude')
        ax1.set_ylim(0, 1.5)
        ax3 = ax1.text(0.3, 0.9, 'SINAD', transform=ax1.transAxes)
        line1, = ax1.plot(dict['Freq'], dict['Freq_level']) # frequency VS level plot

        fig.subplots_adjust(hspace=0.6)

        while True:
            dict = self.get_sample()
            line0.set_ydata(dict['Time_level']) # update data in plots only in loop instead of update entire plots
            line1.set_ydata(dict['Freq_level'])
            ax3.set_text('SINAD = %.2f dB' % self.get_SINAD(dict=dict)) # update SINAD value

            plt.pause(0.01)# this command will call plt.show()

class SpecAn(object):
    def __init__(self, address):
        self.rm = visa.ResourceManager()
        self.SP = self.rm.open_resource(address) # Spec An
        self.book = load_workbook(filename = "Test_Setup.xlsx") # load an existing .xlsx file

    def FEP_Setup(self, freq):
        sheet = self.book["Freq_Error_Power"] # load excel sheet
        # below codes are for setting test frequency in Test_Setup.xlsx according to user's input
        sheet.cell(row = 1, column = 2, value = freq) # write test frequency
        self.book.save("Test_Setup.xlsx") # save existing .xlsx file
        # above codes are for setting test frequency in Test_Setup.xlsx according to user's input

        Centre_frequency = sheet["B1"].value # get all parameters from Test_Setup.xlsx
        Span_frequency = sheet["B2"].value
        RBW = sheet["B3"].value
        VBW = sheet["B4"].value
        RF_level = sheet["B5"].value
        Attenuation = sheet["B6"].value
        RefLev_offset = sheet["B7"].value
        Trace_peak = sheet["B8"].value
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

        self.SP.write(f"*RST") # write all paramaters to SpecAn
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write(f"FREQ:CENT {Centre_frequency}MHz")
        self.SP.write(f"FREQ:SPAN {Span_frequency}Hz")
        self.SP.write(f"BAND {RBW}Hz")
        self.SP.write(f"BAND:VID {VBW}Hz")
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {RF_level}")
        self.SP.write(f"INP:ATT {Attenuation}")
        self.SP.write(f"{RefLev_offset}")
        self.SP.write(f"{Trace_peak}")
        self.SP.write(f"CORR:TRAN:SEL '{Transducer1}'")
        self.SP.write(f"CORR:TRAN {Trans1_ON} ")
        self.SP.write(f"CORR:TRAN:SEL '{Transducer2}'")
        self.SP.write(f"CORR:TRAN {Trans2_ON}")
        self.SP.write(f"CORR:TRAN:SEL '{Transducer3}'")
        self.SP.write(f"CORR:TRAN {Trans3_ON}")
        self.SP.write(f"CALC:LIM:NAME '{Limit_line_1}'")
        self.SP.write(f"CALC:LIM:UPP:STAT {Limit_line_1_ON}")
        self.SP.write(f"CALC:LIM:NAME '{Limit_line_2}'")
        self.SP.write(f"CALC:LIM:UPP:STAT {Limit_line_2_ON}")
        self.SP.write(f"SWE:POIN {Sweep_points}")


    def DeMod_Setup(self, freq):
        sheet = self.book["Analog_Demod"]
        sheet.cell(row = 1, column = 2, value = freq)
        self.book.save("Test_Setup.xlsx")

        Centre_frequency = sheet["B1"].value #get all parameters from Test_Setup.xlsx
        Dev_PerDivision = sheet["B2"].value
        Demod_BW = sheet["B3"].value
        AF_Couple = sheet["B4"].value
        RF_level = sheet["B5"].value
        Attenuation = sheet["B6"].value
        RefLev_offset = sheet["B7"].value
        Trace_Peak = sheet["B8"].value
        Demod_MT = sheet["B9"].value
        Cont_sweep = sheet["B10"].value

        self.SP.write(f"*RST") # write all paramaters to SpecAn
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write("ADEM ON")
        self.SP.write(f"FREQ:CENT {Centre_frequency}MHz")
        self.SP.write(f"DISP:TRAC:Y:PDIV {Dev_PerDivision}kHz")
        self.SP.write(f"BAND:DEM {Demod_BW}kHz")
        self.SP.write(f"ADEM:AF:COUP {AF_Couple}")#Set AF coupling to AC, the frequency offset is automatically corrected.
        # i.e. the trace is always symmetric with respect to the zero line
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {RF_level}")
        self.SP.write(f"INP:ATT {Attenuation}")
        self.SP.write(f"{Trace_Peak}")
        self.SP.write(f"ADEM:MTIM {Demod_MT}ms")
        self.SP.write(f"INIT:CONT {Cont_sweep}")


    def ACP_Setup(self, freq):
        sheet = self.book["ACP"]
        sheet.cell(row = 1, column = 2, value = freq)
        self.book.save("Test_Setup.xlsx")

        Centre_frequency = sheet["B1"].value # get all parameters from Test_Setup.xlsx
        Span_frequency = sheet["B2"].value
        RBW = sheet["B3"].value
        VBW = sheet["B4"].value
        RF_level = sheet["B5"].value
        Attenuation = sheet["B6"].value
        RefLev_offset = sheet["B7"].value
        Trace_RMS = sheet["B8"].value
        Tx_CHBW = sheet["B10"].value
        AJ_CHBW = sheet["B11"].value
        AT_CHBW = sheet["B12"].value
        AJ_CHNUM = sheet["B13"].value
        AJ_SPACE = sheet["B14"].value
        AT_SPACE = sheet["B15"].value
        Power_Mode = sheet["B16"].value
        Ave_number = sheet["B17"].value

        self.SP.write(f"*RST") # write all paramaters to SpecAn
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write("CALC:MARK:FUNC:POW:SEL ACP")
        self.SP.write(f"FREQ:CENT {Centre_frequency}MHz")
        self.SP.write(f"FREQ:SPAN {Span_frequency}kHz")
        self.SP.write(f"BAND {RBW}Hz")
        self.SP.write(f"BAND:VID {VBW}Hz")
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {RF_level}")
        self.SP.write(f"INP:ATT {Attenuation}")
        self.SP.write(f"{Trace_RMS}")
        self.SP.write(f"POW:ACH:BWID:CHAN1 {Tx_CHBW}kHz")
        self.SP.write(f"POW:ACH:BWID:ACH {AJ_CHBW}kHz")
        self.SP.write(f"POW:ACH:BWID:ALT1 {AT_CHBW}kHz")
        self.SP.write(f"POW:ACH:ACP {AJ_CHNUM}")
        self.SP.write(f"POW:ACH:SPAC {AJ_SPACE}kHz")
        self.SP.write(f"POW:ACH:SPAC:ALT1 {AT_SPACE}kHz")
        self.SP.write(f"POW:ACH:MODE {Power_Mode}")
        self.SP.write(f"SWE:COUN {Ave_number}")
        self.SP.write(f"CALC:MARK:FUNC:POW:MODE WRIT")
        self.SP.write(f"DISP:TRAC:MODE AVER")


    def CSE_Setup(self, sub_range, limit_line):
        if sub_range == 1:# choose sub_range
            sheet = self.book["Cond_Spurious_1"]
        elif sub_range == 2:
            sheet = self.book["Cond_Spurious_2"]
        elif sub_range == 3:
            sheet = self.book["Cond_Spurious_3"]
        elif sub_range == 4:
            sheet = self.book["Cond_Spurious_4"]
        elif sub_range == 5:
            sheet = self.book["Cond_Spurious_5"]
        else:
            raise ValueError('sub_range only accept integer 1 or 5')

        if limit_line == 2:
            sheet.cell(row = 16, column = 2, value = "ON") # choose limit line
            sheet.cell(row = 18, column = 2, value = "OFF")
            self.book.save("Test_Setup.xlsx")
        elif limit_line == 3:
            sheet.cell(row = 16, column = 2, value = "OFF") # choose limit line
            sheet.cell(row = 18, column = 2, value = "ON")
            self.book.save("Test_Setup.xlsx")
        else:
            raise ValueError('limit_line only accept integer 2 or 3')

        start_frequency = sheet["B1"].value # get all parameters from Test_Setup.xlsx
        stop_frequency = sheet["B2"].value
        RBW = sheet["B3"].value
        VBW = sheet["B4"].value
        RF_level = sheet["B5"].value
        Attenuation = sheet["B6"].value
        RefLev_offset = sheet["B7"].value
        Trace_peak = sheet["B8"].value
        Transducer1 = sheet["B9"].value
        Trans1_ON = sheet["B10"].value
        Transducer2 = sheet["B11"].value
        Trans2_ON = sheet["B12"].value
        Transducer3 = sheet["B13"].value
        Trans3_ON = sheet["B14"].value
        Limit_line_1 = sheet["B15"].value
        Limit_line_1_ON = sheet["B16"].value
        Limit_line_2 = sheet["B17"].value
        Limit_line_2_ON = sheet["B18"].value
        Sweep_points= sheet["B19"].value

        self.SP.write(f"*RST") # write all paramaters to SpecAn
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write(f"FREQ:STAR {start_frequency}MHz")
        self.SP.write(f"FREQ:STOP {stop_frequency}MHz")
        self.SP.write(f"BAND {RBW}kHz")
        self.SP.write(f"BAND:VID {VBW}kHz")
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {RF_level}")
        self.SP.write(f"INP:ATT {Attenuation}")
        self.SP.write(f"{RefLev_offset}")
        self.SP.write(f"{Trace_peak}")
        self.SP.write(f"CORR:TRAN:SEL '{Transducer1}'")
        self.SP.write(f"CORR:TRAN {Trans1_ON} ")
        self.SP.write(f"CORR:TRAN:SEL '{Transducer2}'")
        self.SP.write(f"CORR:TRAN {Trans2_ON}")
        self.SP.write(f"CORR:TRAN:SEL '{Transducer3}'")
        self.SP.write(f"CORR:TRAN {Trans3_ON}")
        self.SP.write(f"CALC:LIM:NAME '{Limit_line_1}'")
        self.SP.write(f"CALC:LIM:UPP:STAT {Limit_line_1_ON}")
        self.SP.write(f"CALC:LIM:NAME '{Limit_line_2}'")
        self.SP.write(f"CALC:LIM:UPP:STAT {Limit_line_2_ON}")
        self.SP.write(f"SWE:POIN {Sweep_points}")
        self.SP.write(f"DISP:TRAC:MODE MAXH")

    def TranP_Setup(self, freq):
        sheet = self.book["Tran_Perform"]
        sheet.cell(row = 1, column = 2, value = freq)
        self.book.save("Test_Setup.xlsx")

        Centre_frequency = sheet["B1"].value # get all parameters from Test_Setup.xlsx
        Dev_PerDivision = sheet["B2"].value
        Demod_BW = sheet["B3"].value
        AF_Couple = sheet["B4"].value
        RF_level = sheet["B5"].value
        Attenuation = sheet["B6"].value
        RefLev_offset = sheet["B7"].value
        Trace_Peak = sheet["B8"].value
        Demod_MT = sheet["B9"].value
        Cont_sweep = sheet["B10"].value
        Trigger_offset = sheet["B11"].value

        self.SP.write(f"*RST") # write all paramaters to SpecAn
        self.SP.write("SYST:DISP:UPD ON")
        self.SP.write("ADEM ON")
        self.SP.write(f"FREQ:CENT {Centre_frequency}MHz")
        self.SP.write(f"DISP:TRAC:Y:PDIV {Dev_PerDivision}kHz")
        self.SP.write(f"BAND:DEM {Demod_BW}kHz")
        self.SP.write(f"ADEM:AF:COUP {AF_Couple}")
        self.SP.write(f"DISP:TRAC:Y:RLEV:OFFS {RefLev_offset}")
        self.SP.write(f"DISP:TRAC:Y:RLEV {RF_level}")
        self.SP.write(f"INP:ATT {Attenuation}")
        self.SP.write(f"{Trace_Peak}")
        self.SP.write(f"ADEM:MTIM {Demod_MT}ms")
        self.SP.write(f"INIT:CONT {Cont_sweep}")
        self.SP.write(f"TRIG:HOLD {Trigger_offset}")

    def write(self, str):
        self.SP.write(str)

    def query(self, str):
        return self.SP.query(str)

    def get_FEP_result(self):
        self.SP.write("CALC:MARK1:MAX")
        Frequency = self.SP.query("CALC:MARK1:X?")
        Level = float(self.SP.query("CALC:MARK1:Y?"))
        Frequency_error = float(Frequency) - float(self.SP.query("FREQ:CENT?"))*1e6
        indication = (self.SP.query("*OPC?")).replace("1","Completed.")
        return {'F':Frequency_error, 'P':Level, 'I':indication}

    def get_CSE_result(self):
        self.SP.write("CALC:MARK1:MAX")
        Frequency = float(self.SP.query("CALC:MARK1:X?"))/1e6
        Level = float(self.SP.query("CALC:MARK1:Y?"))
        indication = (self.SP.query("*OPC?")).replace("1","Completed.")
        return {'F':Frequency, 'P':Level, 'I':indication}

    def screenshot(self, file_name):
        self.SP.write("HCOP:DEV:LANG PNG") # set file type to .png
        self.SP.write("HCOP:CMAP:DEF4")
        self.SP.write(f"MMEM:NAME \'c:\\temp\\Dev_Screenshot.png\'")
        self.SP.write("HCOP:IMM") # perform copy and save .png file on SpecAn
        self.SP.query("*OPC?")

        file_data = self.SP.query_binary_values(f"MMEM:DATA? \'c:\\temp\\Dev_Screenshot.png\'", datatype='s',)[0] # query binary data and save
        new_file = open(f"c:\\Temp\\{file_name}.png", "wb")# extract file_name string using (f"{}")
        new_file.write(file_data) # copy data to the file on PC
        new_file.close()
        print(f"saved to PC c:\\Temp\\{file_name}.png\n")

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

    def Tx_Setup(self):
        self.sheet = self.book["ACP"]
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
    FSV.DeMod_Setup(460)
    FSV.query('*OPC?')

    CP50.Set_Freq(460)
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

    FSV.ACP_Setup(460)
    time.sleep(8)
    FSV.write(f"DISP:TRAC:MODE VIEW")


    ACP = FSV.query("CALC:MARK:FUNC:POW:RES? ACP")
    print(ACP)
    LIST = re.findall(r'\d+\.\d+', ACP)
    Result_sheet.write(row = 2, column = 1, value = FSV.query("FREQ:CENT?"))
    Result_sheet.write(row = 2, column = 2, value = float(LIST[0]))
    Result_sheet.write(row = 2, column = 3, value = -float(LIST[1]))
    Result_sheet.write(row = 2, column = 4, value = -float(LIST[2]))

    Result_sheet.save("Test_Result.xlsx") # save existing .xlsx file
    CP50.Radio_Off()
    SML.write(":OUTP2 OFF")# turn off audio output at the end of the test
    indication = (FSV.query("*OPC?")).replace("1","Completed.")
    print(f"ACP test {indication}")


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

    Start_F = Decimal(479.9875)/Decimal(1) # Decimal module make sure float numbers addition yeilds correct value
    CP50_Result.get_sheet("day2")

    for i in range(9, 11):
        Spur_Res = Rx_Spurious_response_immunity(freq=Start_F, delta=Decimal(-2*38.85))
        CP50_Result.write(row = i+2, column = 1, value = Start_F)
        CP50_Result.write(row = i+2, column = 2, value = Start_F+Decimal(-2*38.85))
        CP50_Result.write(row = i+2, column = 3, value = Spur_Res)
        Timestamp ='{:%d-%b-%Y %H:%M:%S}'.format(datetime.datetime.now())
        CP50_Result.write(row = i+2, column = 4, value = Timestamp)
        CP50_Result.save("CP50_Result.xlsx")
        Start_F += Decimal(0.0125) # frequency in MHz


try:
    SC = SoundCard()
except BaseException:
    print("Specified Soundcard does not exist.")
    pass

try:
    FSV = SpecAn('TCPIP0::192.168.10.9::hislip0::INSTR')
except BaseException:
    print("FSV is not on.")
    pass

try:
    CP50 = Radio('com7')
except BaseException:
    print("Specified com port does not exsit.")
    pass

try:
    SML = SigGen('ASRL4::INSTR')
except BaseException:
    print("SML is not on.")
    pass

try:
    SMB = SigGen('USB0::0x0AAD::0x0054::106409::INSTR')
except BaseException:
    print("SMB is not on.")
    pass

try:
    CMS = SigGen('GPIB0::24::INSTR')
except BaseException:
    print("CMS is not on.")
    pass

try:
    Result_sheet = Excel("Test_Result.xlsx")
    CP50_Result = Excel("CP50_Result.xlsx")
except BaseException:
    print("Specified Excel file does not exsit.")
    pass

#SC.live_plots()
#Tx_Adjacent_channel_power()
#Tx_Frequency_error_Carrier_power()
#Tx_Conducted_spurious_emissions()

getcontext().prec = 10 # set 10 decimal values precision
