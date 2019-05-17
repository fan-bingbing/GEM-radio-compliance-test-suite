import sys
import math


# below codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\midlevel')
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\bottomlevel')
# above codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder

import config # containing global variables
import Rx_Write_to_excel as RxW
import FSV_screenshot as Scr
import RadioControl as Rcon
import FreqError_Power as FEP
import Max_Deviation as MAD
import OOB_ModRes as OOB
import ACP
import Cond_Spur as Cons
import ACS
import Blocking as BLK


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
    print(ACP.Adjacent_channel_power(Test_frequency))
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
    print("--------------------------------------------")
    print("Conducted_spurious_emissions_test_menu")
    print("--------------------------------------------")
    print("1. 9k-150kHz")
    print("2. 150kHz-30MHz")
    print("3. 30MHz-700MHz")
    print("4. 700MHz-1GHz")
    print("5. 1-4GHz")
    print("0. Exit")
    while True:
        choice = input("> ")
        if choice == "1":
            print("Check physical test setup: 30dB attenuator in line WITHOUT high pass filter ?")
            dict = Conducted_spurious_menu("Cond_Spurious_1")
            print(f"Mark frequency:{dict['Mark frequency']/1e3}kHz")
            print(f"Mark level:{dict['Mark level']}dBm")
            break
        elif choice == "2":
            print("Check physical test setup: 30dB attenuator in line WITHOUT high pass filter ?")
            dict = Conducted_spurious_menu("Cond_Spurious_2")
            print(f"Mark frequency:{dict['Mark frequency']/1e6}MHz")
            print(f"Mark level:{dict['Mark level']}dBm")
            break
        elif choice == "3":
            print("Check physical test setup: 30dB attenuator in line WTIHOUT high pass filter ?")
            dict = Conducted_spurious_menu("Cond_Spurious_3")
            print(f"Mark frequency:{dict['Mark frequency']/1e6}MHz")
            print(f"Mark level:{dict['Mark level']}dBm")
            break
        elif choice == "4":
            print("Check physical test setup: 30dB attenuator in line WITH high pass filter NHP700+ ?")
            dict = Conducted_spurious_menu("Cond_Spurious_4")
            print(f"Mark frequency:{dict['Mark frequency']/1e6}MHz")
            print(f"Mark level:{dict['Mark level']}dBm")
            break
        elif choice == "5":
            print("Check physical test setup: 30dB attenuator in line WITH high pass filter NHP700+ ?")
            dict = Conducted_spurious_menu("Cond_Spurious_5")
            print(f"Mark frequency:{dict['Mark frequency']/1e6}MHz")
            print(f"Mark level:{dict['Mark level']}dBm")
            break

        elif choice == "0":
            exit(0)
        else:
            print("please enter number within the range from 0 to 5.")

    Tx_Conducted_spurious_emissions() # go back to conducted spurious menu after break



def Conducted_spurious_menu(frequency_range):

    print("1. YES, Continue")
    print("2. NO, Go back to previous menu")
    while True:
        choice = input("> ")
        if choice == "1":
            Rcon.RadioControl('com7', 'ON')
            dict = Cons.Conducted_spurious(frequency_range)
            print(f"Conducted_spurious_emissions_test {dict['Indication']}")
            #print(f"Mark frequency:{dict['Mark frequency']}Hz")
            #print(f"Mark level:{dict['Mark level']}dBm")
            Rcon.RadioControl('com7', 'OFF')
            file_name = input("To save the screenshot, input the filename (***.bmp) or press CTRL+C to quit > ")
            Scr.Screenshot(file_name)
            return dict
            break

        elif choice == "2":
            Tx_Conducted_spurious_emissions()
        else:
            print("please enter number within the range from 1 or 2.")


def Rx_Spurious_emissions():
    print("blank")
    exit(0)

def Rx_Adjacent_channel_selectivity():
    Test_frequency = input("Input test frequency in Mhz or press CTRL+C to quit > ")
    dict = ACS.Adjacent_channel_selectivity(Test_frequency)
    print (dict['Indication'])
    print (f"ACS+: {dict['ACS_high']}, ACS-: {dict['ACS_low']}")
    exit(0)

def Rx_Spurious_response_immunity():
    print("blank")
    exit(0)


def Rx_Blocking_immunity():
    Test_frequency = input("Input test frequency in Mhz or press CTRL+C to quit > ")
    dict = BLK.Blocking_immunity(Test_frequency)
    print (dict['Indication'])
    print (f"BLK+: {dict['BLK_high']}, BLK-: {dict['BLK_low']}")
    exit(0)
