import Modules
from sys import exit

header = """
GME Radio compliance test suite Version 1.0
"""

print("--------------------------------------------")
print(header)
print("--------------------------------------------")



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
    print("1. ASNZS4365 for CB Radio")
    print("2. ASNZS4295 for Comercial Radio")
    print("3. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            ANZ4365()
        elif choice == "2":
            ANZ4295()
        elif choice == "3":
            exit(0)
        else:
            print("please enter number 1, 2, or 3.")

start()
