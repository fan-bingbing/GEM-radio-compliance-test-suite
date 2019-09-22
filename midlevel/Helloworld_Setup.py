import visa
rm = visa.ResourceManager()


#scope.write_termination = '\n'

def Hello_world():

    print(rm.list_resources())

    SML = rm.open_resource('ASRL4::INSTR') # Unwanted Signal
    SMB = rm.open_resource('USB0::0x0AAD::0x0054::106409::INSTR') # Wanted Signal
    FSV = rm.open_resource('TCPIP0::192.168.10.9::hislip0::INSTR') # Spec An
    CMS = rm.open_resource('GPIB0::24::INSTR') # Audio Analyzer
    PS = rm.open_resource('GPIB0::8::INSTR') # Power supply

    SML.clear()  # Clear instrument io buffers and status
    SMB.clear()
    FSV.clear()
    CMS.clear()
    PS.clear()

    idn_response1 = SML.query('*IDN?')  # Query the Identification string
    idn_response2 = SMB.query('*IDN?')
    idn_response3 = FSV.query('*IDN?')
    idn_response4 = CMS.query('*IDN?')
    idn_response5 = PS.query('*IDN?')

    print (f"Hello, I am {idn_response1}")
    print (f"Hello, I am {idn_response2}")
    print (f"Hello, I am {idn_response3}")
    print (f"Hello, I am {idn_response4}")
    print (f"Hello, I am {idn_response5}")

    SML.close()  # Clear instrument io buffers and status
    SMB.close()
    FSV.close()
    CMS.close()
    PS.close()
