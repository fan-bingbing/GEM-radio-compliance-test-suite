import serial
import time

ESC_CHAR = 0x7D # Escape char
STX_CHAR =  0x7E # Start of packet
ETX_CHAR = 0x7F # End of packet
CHECKSUM_XOR_MASK = 0xFF


def Packet_Gen(payload):
    # Header
    tx_array = bytearray((STX_CHAR, 2, len(payload)))

    # Escape the payload, append it to the output buffer, make the checksum
    chksum = 0
    for i, b in enumerate(payload):
        # The first word of the payload is always the command ID.
        # The MSB of the command ID is always 1. The radio just
        # clears it in the ACK, without doing anything else with it.
        if i == 0:
            bb = (0x80 | b) & 0xFF
        else:
            bb = b & 0xFF
        chksum ^= bb
        if bb in (ESC_CHAR, STX_CHAR, ETX_CHAR):
            tx_array.append(ESC_CHAR)
            tx_array.append(bb ^ 0x20)
        else:
            tx_array.append(bb)

    # Trailer
    tx_array.append(chksum ^ CHECKSUM_XOR_MASK)
    tx_array.append(ETX_CHAR)

    return tx_array


def Radio_Con(port, baudrate, timeout, payload):
    ser = serial.Serial(port=port, baudrate=baudrate, timeout=timeout)  # open serial port
    time.sleep(0.5)
    ser.write(Packet_Gen(payload))
    print(ser.name) # set a indicator showing one command has been sent to radio
    ser.close()
