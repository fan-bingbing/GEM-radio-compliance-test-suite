import sys
import math
# below codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\midlevel')
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\bottomlevel')
# above codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder

import FSV_screenshot as Scr
import RadioControl as Rcon
import FreqError_Power as FEP
import Max_Deviation as MAD
import OOB_ModRes as OOB
import ACP


def Tx_Frequency_error_Carrier_power():
    Test_frequency = input("Input test frequency in Mhz or press CTRL+C to quit > ")
    Rcon.RadioControl('com7', 'ON')
    dict = FEP.FreqError_Power(Test_frequency)
    print(f"Frequency error and Carrier power test {dict['Indication']}")
    print(f"Frequency error:{dict['Frequency_error']}Hz")
    print(f"Carrier power:{dict['Carrier_power']}dBm")
    Rcon.RadioControl('com7', 'OFF')
    file_name = input("To save the screenshot, input the filename (***.bmp) or press CTRL+C to quit > ")
    Scr.Screenshot(file_name)
    exit(0)

def Tx_Max_deviation():
    Test_frequency = input("Input test frequency in Mhz or press CTRL+C to quit > ")
    Rcon.RadioControl('com7', 'ON')
    list = MAD.Max_Deviation(Test_frequency)
    print(f"Max deviation test {list[13]}")
    print(f"Result stored in:{list}")
    Rcon.RadioControl('com7', 'OFF')
    exit(0)

def Tx_Adjacent_channel_power():
    Test_frequency = input("Input test frequency in Mhz or press CTRL+C to quit > ")
    Rcon.RadioControl('com7', 'ON')
    print(ACP.Tx_Adjacent_channel_power(Test_frequency))
    Rcon.RadioControl('com7', 'OFF')
    file_name = input("To save the screenshot, input the filename (***.bmp) or press CTRL+C to quit > ")
    Scr.Screenshot(file_name)
    exit(0)

def Tx_Transient_performance():
    print("blank")
    exit(0)

def Tx_Out_of_band_modulation_response():
    Test_frequency = input("Input test frequency in Mhz or press CTRL+C to quit > ")
    Rcon.RadioControl('com7', 'ON')
    Abs_deviation = OOB.OOB_ModRes(Test_frequency)
    Rel_deviation = []
    for i in range(0,8):
        Rel_deviation.append(20*math.log10(Abs_deviation[i]/Abs_deviation[8]))
    print(f"Out of Band Modulation Response test {Abs_deviation[9]}")
    print(f"Absolution deviation:{Abs_deviation}")
    print(f"Relative deviation:{Rel_deviation}")
    Rcon.RadioControl('com7', 'OFF')

    exit(0)

def Tx_Conducted_spurious_emissions():
    print("blank")
    exit(0)

def Rx_Spurious_emissions():
    print("blank")
    exit(0)

def Rx_Adjacent_channel_selectivity():
    print("blank")
    exit(0)

def Rx_Spurious_response_immunity():
    print("blank")
    exit(0)


def Rx_Blocking_immunity():
    print("blank")
    exit(0)
