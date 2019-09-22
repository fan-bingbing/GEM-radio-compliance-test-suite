"""
GME radio compliance test suite scripts for toplevel menu .
Copyright (C) 2019 Standard Communications Pty Ltd (GME). All rights reserved.
"""



import sys
import Module
import config
print(sys.executable)
print(sys.version)


header = """
GME Radio compliance test suite Version 1.0
"""

def Helloworld():

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
    print("7. Run all Tx test on standard required channels")
    print("8. Rx-Spurious emissions")
    print("9. Go back to last menu")
    print("0. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.CHSW_FEP(config.test_frequency, start_row=0, end_row=1, clear=1)
            ANZ4365()
        elif choice == "2":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.Tx_Max_deviation_CB(config.test_frequency, sheet="Max_Dev_CB")
            ANZ4365()
        elif choice == "3":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.CHSW_ACP(config.test_frequency, start_row=0, end_row=1, clear=1)
            ANZ4365()
        elif choice == "4":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.Tx_Transient_performance(config.test_frequency)
            ANZ4365()
        elif choice == "5":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.Tx_Out_of_band_modulation_response(config.test_frequency, sheet="OOB_CB")
            ANZ4365()
        elif choice == "6":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.Tx_Conducted_spurious_emissions(config.test_frequency, sheet="Cond_Spur_CB", clear=0)
            ANZ4365()
        elif choice == "7":
            Module.SW1.Switch_to_Ax(1)
            Module.CHSW_FEP(476.9, start_row=0, end_row=1, clear=1)# clear sheet in first run
            Module.Tx_Transient_performance(476.9)
            Module.Tx_Out_of_band_modulation_response(476.9, sheet="OOB_CB")
            Module.Tx_Max_deviation_CB(476.9, sheet="Max_Dev_CB")
            Module.CHSW_ACP(476.9, start_row=0, end_row=1, clear=1)# clear sheet in first run
            ANZ4365()
        elif choice == "8":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.Rx_Spurious_emissions(config.test_frequency)
            ANZ4365()
        elif choice == "9":
            start()
        elif choice == "0":
            Module.FSV.close()
            Module.SML.close()
            Module.SMB.close()
            Module.SMC.close()
            Module.CMS.close()
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
    print("8. Run all Tx test on standard required channels")
    print("9. Run all Rx test on standard required channels")
    print("10. Go back to last menu")
    print("0. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.CHSW_FEP(config.test_frequency, start_row=0, end_row=1, clear=1)
            ANZ4295()
        elif choice == "2":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.Tx_Max_deviation(config.test_frequency, sheet="Max_Dev_1")
            ANZ4295()
        elif choice == "3":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(1)
            Module.CHSW_ACP(config.test_frequency, start_row=0, end_row=1, clear=1)
            ANZ4295()
        elif choice == "4":
            config.test_frequency = input("input test frequency in MHz > ")
            saved_sheet = input("input sheet the result to be saved\n for example: Cond_Spur_1/2/3 > ")
            Module.SW1.Switch_to_Ax(1)
            Module.Tx_Conducted_spurious_emissions(config.test_frequency, sheet=str(saved_sheet), clear=0)
            ANZ4295()
        elif choice == "5":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(2)
            Module.CHSW_ACS(config.test_frequency, start_row=0, end_row=1, clear=1)
            ANZ4295()
        elif choice == "6":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(2)
            Module.CHSW_SR(config.test_frequency, start_row=0, end_row=1, clear=1)
            ANZ4295()
        elif choice == "7":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(2)
            Module.CHSW_BLK(config.test_frequency, start_row=0, end_row=1, clear=1)
            ANZ4295()
        elif choice == "8":
            Module.SW1.Switch_to_Ax(1)
            Module.CHSW_FEP(459.075, start_row=0, end_row=1, clear=1)# clear sheet in first run
            Module.CHSW_FEP(498.700, start_row=1, end_row=2, clear=0)
            Module.CHSW_FEP(519.200, start_row=2, end_row=3, clear=0)
            Module.Tx_Max_deviation(459.075, sheet="Max_Dev_1")
            Module.Tx_Max_deviation(498.700, sheet="Max_Dev_2")
            Module.Tx_Max_deviation(519.200, sheet="Max_Dev_3")
            Module.CHSW_ACP(459.075, start_row=0, end_row=1, clear=1)# clear sheet in first run
            Module.CHSW_ACP(498.700, start_row=1, end_row=2, clear=0)
            Module.CHSW_ACP(519.200, start_row=2, end_row=3, clear=0)
            ANZ4295()
        elif choice == "9":
            Module.SW1.Switch_to_Ax(2)
            Module.CHSW_ACS(459.075, start_row=0, end_row=1, clear=1)# clear sheet in first run
            Module.CHSW_ACS(498.700, start_row=1, end_row=2, clear=0)
            Module.CHSW_ACS(519.200, start_row=2, end_row=3, clear=0)
            Module.CHSW_SR(459.075, start_row=0, end_row=1, clear=1)
            Module.CHSW_SR(498.700, start_row=1, end_row=2, clear=0)
            Module.CHSW_SR(519.200, start_row=2, end_row=3, clear=0)
            Module.CHSW_BLK(459.075, start_row=0, end_row=1, clear=1)
            Module.CHSW_BLK(498.700, start_row=1, end_row=2, clear=0)
            Module.CHSW_BLK(519.200, start_row=2, end_row=3, clear=0)
            ANZ4295()
        elif choice == "10":
            start()
        elif choice == "0":
            Module.FSV.close()
            Module.SML.close()
            Module.SMB.close()
            Module.SMC.close()
            Module.CMS.close()
            exit(0)
        else:
            print("please enter number within the range from 0 to 8.")

def EN300086():
    print("--------------------------------------------")
    print("EN300086 V2.1.2")
    print("--------------------------------------------")
    print("1. Rx-Maximum usable sensitivity(conducted)")
    print("2. Rx-Co-channel rejection")
    print("3. Rx-Intermodulation response rejection")
    print("4. Run all Rx test on standard required channels")
    print("5. Go back to last menu")
    print("0. Quit")
    while True:
        choice = input("> ")
        if choice == "1":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(2)
            Module.CHSW_MUS(config.test_frequency, start_row=0, end_row=1, clear=1)
            EN300086()
        elif choice == "2":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(2)
            Module.CHSW_CCR(config.test_frequency, start_row=0, end_row=1, clear=1)
            EN300086()
        elif choice == "3":
            config.test_frequency = input("input test frequency in MHz > ")
            Module.SW1.Switch_to_Ax(2)
            Module.CHSW_Intermodulation(config.test_frequency, start_row=0, end_row=1, clear=1)
            EN300086()
        elif choice == "4":
            Module.SW1.Switch_to_Ax(2)
            Module.CHSW_MUS(451.0125, start_row=0, end_row=1, clear=1) # clear sheet in first run
            Module.CHSW_MUS(518.9875, start_row=2, end_row=3, clear=0)
            Module.CHSW_CCR(451.0125, start_row=0, end_row=1, clear=1)
            Module.CHSW_CCR(518.9875, start_row=2, end_row=3, clear=0)
            Module.CHSW_Intermodulation(451.0125, start_row=0, end_row=1, clear=1)
            Module.CHSW_Intermodulation(518.9875, start_row=2, end_row=3, clear=0)
            start()
        elif choice == "5":
            start()
        elif choice == "0":
            Module.FSV.close()
            Module.SML.close()
            Module.SMB.close()
            Module.SMC.close()
            Module.CMS.close()
            exit(0)
        else:
            print("please enter number within the range from 0 to 4.")

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
    print("3. Occupied_bandwidth")
    print("4. Screenshot")
    print("5. Tx Result transfer")
    print("6. Rx Result transfer")
    print("7. Go back to last menu")
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
            Module.Tx_Occupied_bandwidth(472.025, "top")
            Module.Tx_Occupied_bandwidth(472.1, "bottom")
        elif choice == "4":
            filename = input("input filename > ")
            Module.FSV.screenshot(filename)
            Misc()
        elif choice == "5":
            filename = input("input filename where Tx result will be saved: > ")
            Module.copy_excel(to_file=filename,
                              sheet_list=['Ferror_Pow', 'ACP', 'Cond_Spur_CB',
                                          'Cond_Spur_3','Cond_Spur_2', 'Cond_Spur_1',
                                          'Max_Dev_3','Max_Dev_2', 'Max_Dev_1',
                                          'OOB_CB','Max_Dev_CB'
                                          ]
                              )

            Misc()

        elif choice == "6":
            filename = input("input filename where Rx result will be saved: > ")
            Module.copy_excel(to_file=filename,
                              sheet_list=['ACS', 'CCR', 'MUS',
                                          'Blocking', 'Spur_Res',
                                          'Intermodulation_Response'
                                          ]
                              )
            Misc()

        elif choice == "7":
            start()
        elif choice == "0":
            Module.FSV.close()
            Module.SML.close()
            Module.SMB.close()
            Module.SMC.close()
            Module.CMS.close()
            exit(0)
        else:
            print("please enter number within the range from 0 to 2.")


def start():
    print("--------------------------------------------")
    print(header)
    print("--------------------------------------------")
    print("1. Hello World!")
    print("2. ASNZS4365 for CB Radio")
    print("3. ASNZS4295 for Commercial Radio")
    print("4. EN300086 for LandMobile Radio")
    print("5. AS/NZS ETSI EN301 178 for VHF maritime radio(non-GMDSS appliacation only)")


    print("6. Misc")
    print("0. Quit")

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
            EN300086()
        elif choice == "5":
            ANZS_ETSI_EN301_178()
        elif choice == "6":
            config.limit_line_factor = 4 # assign limit line value info for conducted spurious test
            Misc()
        elif choice == "0":
            Module.FSV.close()
            Module.SML.close()
            Module.SMB.close()
            Module.SMC.close()
            Module.CMS.close()
            exit(0)
        else:
            print("please enter number 0 to 5.")

if __name__ == '__main__':
    start()
