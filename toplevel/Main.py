import sys

print(sys.executable)
print(sys.version)
sys.path.insert(0, r'C:\Users\afan\Documents\commercial-radio-test-suite\midlevel')
sys.path.insert(0, r'C:\Users\afan\Documents\commercial-radio-test-suite\bottomlevel')
sys.path.insert(0, r'C:\Users\afan\Documents\commercial-radio-test-suite\toplevel')

import Module
import config



header = """
GME Radio compliance test suite Version 1.0
"""

print("--------------------------------------------")
print(header)
print("--------------------------------------------")

# Test_Setup.xlsx has to be in the same folder where Main.py locate


def Helloworld():
    print("Blank")
    start()

def ANZ4365():
    print("--------------------------------------------")
    print("ASNZS4365:2011")
    print("--------------------------------------------")
    print("1. Tx-Frequency error_Carrier power")
    print("2. Tx-Max deviation")
    print("3. Tx-Adjacent channel power")
    print("4. Tx-Transient performance")
    print("5. Tx-Out of band modulation response")
    print("6. Tx-Conducted spurious emissions ")
    print("7. Rx-Spurious emissions")
    print("8. Go back to last menu")
    print("0. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            Module.Tx_Frequency_error_Carrier_power()
            ANZ4365()
        elif choice == "2":
            Module.Tx_Max_deviation()
            ANZ4365()
        elif choice == "3":
            Module.Tx_Adjacent_channel_power()
            ANZ4365()
        elif choice == "4":
            Module.Tx_Transient_performance()
            ANZ4365()
        elif choice == "5":
            Module.Tx_Out_of_band_modulation_response()
            ANZ4365()
        elif choice == "6":
            Module.Tx_Conducted_spurious_emissions()
            ANZ4365()
        elif choice == "7":
            Module.Rx_Spurious_emissions()
            ANZ4365()
        elif choice == "8":
            start()
        elif choice == "0":
            exit(0)
        else:
            print("please enter number within the range from 0 to 8.")

def ANZ4295():
    print("--------------------------------------------")
    print("ASNZS4295:2015")
    print("--------------------------------------------")
    print("1. Tx-Frequency error_Carrier power")
    print("2. Tx-Max deviation")
    print("3. Tx-Adjacent channel power")
    print("4. Tx-Conducted spurious emissions")
    print("5. Rx-Adjacent channel selectivity")
    print("6. Rx-Spurious response immunity ")
    print("7. Rx-Blocking immunity")
    print("8. Go back to last menu")
    print("0. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            Module.Tx_Frequency_error_Carrier_power()
            ANZ4295()
        elif choice == "2":
            Module.Tx_Max_deviation()
            ANZ4295()
        elif choice == "3":
            Module.Tx_Adjacent_channel_power()
            ANZ4295()
        elif choice == "4":
            Module.Tx_Conducted_spurious_emissions()
            ANZ4295()
        elif choice == "5":
            Module.Rx_Adjacent_channel_selectivity()
            ANZ4295()
        elif choice == "6":
            Module.Rx_Spurious_response_immunity()
            ANZ4295()
        elif choice == "7":
            Module.Rx_Blocking_immunity()
            ANZ4295()
        elif choice == "8":
            start()
        elif choice == "0":
            exit(0)
        else:
            print("please enter number within the range from 0 to 8.")

def ANZS_ETSI_EN301_178():
    print("--------------------------------------------")
    print("AS/NZS ETSI EN301 178:2018")
    print("--------------------------------------------")
    print("1. Tx-Frequency error_Carrier power")
    print("2. Tx-Max deviation")
    print("3. Tx-Max deviation_Above 3kHz")
    print("4. Sensitivity of the modulator, including microphone")
    print("5. Audio frequency response")
    print("6. Audio frequency harmonic distortion of the emission")
    print("7. Tx-Adjacent channel power")
    print("8. Tx-Conducted spurious emissions conveyed to the antenna")
    print("9. Cabinet radiation")
    print("10.Residual modulation of the transmitter  ")
    print("11.Transient frequency behavior of the transmitter")
    print("0. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            exit(0)
        elif choice == "2":
            exit(0)
        elif choice == "3":
            exit(0)
        elif choice == "4":
            exit(0)
        elif choice == "5":
            exit(0)
        elif choice == "6":
            exit(0)
        elif choice == "7":
            exit(0)
        elif choice == "8":
            exit(0)
        elif choice == "9":
            exit(0)
        elif choice == "10":
            exit(0)
        elif choice == "11":
            exit(0)
        elif choice == "0":
            exit(0)
        else:
            print("please enter number within the range from 0 to 11.")

def Misc():
    print("--------------------------------------------")
    print("Misc")
    print("--------------------------------------------")
    print("1. CP50 Spurious Response full scan")
    print("2. CP50 ACS full scan")
    print("3. Go back to last menu")
    print("0. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            Module.CHSW_SR()
            Misc()
        elif choice == "2":
            Module.CHSW_ACS()
            Misc()
        elif choice == "3":
            start()
        elif choice == "0":
            exit(0)
        else:
            print("please enter number within the range from 0 to 2.")


def start():
    print("1. Hello World!")
    print("2. ASNZS4365 for CB Radio")
    print("3. ASNZS4295 for Commercial Radio")
    print("4. AS/NZS ETSI EN301 178 for VHF maritime radio(non-GMDSS appliacation only)")
    print("5. Misc")
    print("6. Quit")

    while True:
        choice = input("> ")
        if choice == "1":
            Helloworld()
        elif choice == "2":
            config.limit_line_factor = 2 # assign limit line value info for conducted spurious test
            ANZ4365()
        elif choice == "3":
            config.limit_line_factor = 3 # assign limit line value info for conducted spurious test
            ANZ4295()
        elif choice == "4":
            ANZS_ETSI_EN301_178()
        elif choice == "5":
            Misc()
        elif choice == "6":
            exit(0)
        else:
            print("please enter number 1, 2, 3 or 4.")

start()
