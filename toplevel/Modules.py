import sys
# below codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\midlevel')
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\bottomlevel')
# above codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder

import ACP
import FSV_screenshot as Scr


def Tx_Frequency_error():
    print("blank")
    exit(0)

def Tx_Carrier_power():
    print("blank")
    exit(0)

def Tx_Max_deviation():
    print("blank")
    exit(0)

def Tx_Adjacent_channel_power():
    print(ACP.Tx_Adjacent_channel_power())
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
