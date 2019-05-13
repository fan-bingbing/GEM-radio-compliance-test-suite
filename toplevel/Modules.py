import sys
# below codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\midlevel')
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\bottomlevel')
# above codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder

import FSV_screenshot as Scr
import RadioControl as Rcon
import FreqError_Power as FEP
import ACP


def Tx_Frequency_error_Carrier_power():
    Test_frequency = input("Input test frequency in Mhz or press CTRL+C to quit > ")
    Rcon.RadioControl('com8', 'ON')
    dict = FEP.FreqError_Power(Test_frequency)
    print(f"Frequency error and Carrier power test {dict['Indication']}")
    print(f"Frequency error:{dict['Frequency_error']}Hz")
    print(f"Carrier power:{dict['Carrier_power']}dBm")
    Rcon.RadioControl('com8', 'OFF')
    file_name = input("To save the screenshot, input the filename (***.bmp) or press CTRL+C to quit > ")
    Scr.Screenshot(file_name)
    exit(0)

def Tx_Max_deviation():
    print("blank")
    exit(0)

def Tx_Adjacent_channel_power():
    Test_frequency = input("Input test frequency in Mhz or press CTRL+C to quit > ")
    print(ACP.Tx_Adjacent_channel_power(Test_frequency))
    file_name = input("To save the screenshot, input the filename (***.bmp) or press CTRL+C to quit > ")
    Scr.Screenshot(file_name)
    exit(0)

def Tx_Transient_performance():
    print("blank")
    exit(0)

def Tx_Out_of_band_modulation_response():
    print("blank")
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
