from sys import exit
import Modules
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
    print("0. Go back to last menu")
    while True:
        choice = input("> ")
        if choice == "1":
            Modules.Tx_Frequency_error_Carrier_power()
        elif choice == "2":
            Modules.Tx_Max_deviation()
        elif choice == "3":
            Modules.Tx_Adjacent_channel_power()
        elif choice == "4":
            Modules.Tx_Transient_performance()
        elif choice == "5":
            Modules.Tx_Out_of_band_modulation_response()
        elif choice == "6":
            Modules.Tx_Conducted_spurious_emissions()
        elif choice == "7":
            Modules.Rx_Spurious_emissions()
        elif choice == "0":
            start()
        else:
            print("please enter number within the range from 0 to 8.")

def ANZ4295():
    print("--------------------------------------------")
    print("ASNZS4295:2015")
    print("--------------------------------------------")
    print("1. Tx-Frequency error")
    print("2. Tx-Max deviation")
    print("3. Tx-Adjacent channel power")
    print("4. Tx-Conducted spurious emissions")
    print("5. Rx-Adjacent channel selectivity")
    print("6. Rx-Spurious response immunity ")
    print("7. Rx-Blocking immunity")
    print("0. Go back to last menu")
    while True:
        choice = input("> ")
        if choice == "1":
            Modules.Tx_Frequency_error_Carrier_power()
        elif choice == "2":
            Modules.Tx_Max_deviation()
        elif choice == "3":
            Modules.Tx_Adjacent_channel_power()
        elif choice == "4":
            Modules.Tx_Conducted_spurious_emissions()
        elif choice == "5":
            Modules.Rx_Adjacent_channel_selectivity()
        elif choice == "6":
            Modules.Rx_Spurious_response_immunity()
        elif choice == "7":
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
            print("please enter number 1, 2, 3 or 4.")

start()
