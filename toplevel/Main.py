import Modules
from sys import exit
import visa
header = """
GME Radio compliance test suite Version 1.0
"""

print("--------------------------------------------")
print(header)
print("--------------------------------------------")

rm = visa.ResourceManager()
SML = rm.open_resource('ASRL4::INSTR') # Unwanted Signal
SMB = rm.open_resource('USB0::0x0AAD::0x0054::106409::INSTR') # Wanted Signal
FSV = rm.open_resource('TCPIP0::192.168.10.9::hislip0::INSTR') # Spec An
CMS = rm.open_resource('GPIB0::24::INSTR') # Audio Analyzer
SML.clear()  # Clear instrument io buffers and status
SMB.clear()
FSV.clear()
CMS.clear()

def Hellowworld():
    idn_response1 = SML.query('*IDN?')  # Query the Identification string
    idn_response2 = SMB.query('*IDN?')
    idn_response3 = FSV.query('*IDN?')
    idn_response4 = CMS.query('*IDN?')
    print (f"Hello, I am {idn_response1}")
    print (f"Hello, I am {idn_response2}")
    print (f"Hello, I am {idn_response3}")
    print (f"Hello, I am {idn_response4}")
    start()

def ANZ4365():
    print("1. Tx-Frequency error")
    print("2. Tx-Carrier power")
    print("3. Tx-Max deviation")
    print("4. Tx-Adjacent channel power")
    print("5. Tx-Transient performance")
    print("6. Tx-Out of band modulation response")
    print("7. Tx-Conducted spurious emissions ")
    print("8. Rx-Spurious emissions")
    print("0. Go back to last menu")
    while True:
        choice = input("> ")
        if choice == "1":
            Modules.Tx_Frequency_error()
        elif choice == "2":
            Modules.Tx_Carrier_power()
        elif choice == "3":
            Modules.Tx_Max_deviation()
        elif choice == "4":
            Modules.Tx_Adjacent_channel_power()
        elif choice == "5":
            Modules.Tx_Transient_performance()
        elif choice == "6":
            Modules.Tx_Out_of_band_modulation_response()
        elif choice == "7":
            Modules.Tx_Conducted_spurious_emissions()
        elif choice == "8":
            Modules.Rx_Spurious_emissions()
        elif choice == "0":
            start()
        else:
            print("please enter number within the range from 0 to 8.")

def ANZ4295():
    print("1. Tx-Frequency error")
    print("2. Tx-Carrier power")
    print("3. Tx-Max deviation")
    print("4. Tx-Adjacent channel power")
    print("5. Tx-Conducted spurious emissions")
    print("6. Rx-Adjacent channel selectivity")
    print("7. Rx-Spurious response immunity ")
    print("8. Rx-Blocking immunity")
    print("0. Go back to last menu")
    while True:
        choice = input("> ")
        if choice == "1":
            Modules.Tx_Frequency_error()
        elif choice == "2":
            Modules.Tx_Carrier_power()
        elif choice == "3":
            Modules.Tx_Max_deviation()
        elif choice == "4":
            Modules.Tx_Adjacent_channel_power()
        elif choice == "5":
            Modules.Tx_Conducted_spurious_emissions()
        elif choice == "6":
            Modules.Rx_Adjacent_channel_selectivity()
        elif choice == "7":
            Modules.Rx_Spurious_response_immunity()
        elif choice == "8":
            Modules.Rx_Blocking_immunity()
        elif choice == "0":
            start()
        else:
            print("please enter number within the range from 0 to 8.")

def start():
    print("1. Hello World!")
    print("2. ASNZS4365 for CB Radio")
    print("3. ASNZS4295 for Commercial Radio")
    print("4. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            Helloworld()
        elif choice == "2":
            ANZ4365()
        elif choice == "3":
            ANZ4295()
        elif choice == "4":
            exit(0)
        else:
            print("please enter number 1, 2, or 3.")

start()
