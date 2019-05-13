import serial
import time

def RadioControl(port, indicator):
    ser = serial.Serial(port=port, baudrate=9600, timeout=None)  # open serial port

    time.sleep(0.5)
    if indicator == 'ON':
        ser.write(b'AT+WGPTT=0,1\n')# write AT command as binary to serial

    elif indicator == 'OFF':
        ser.write(b'AT+WGPTT=0,0\n')

    else:
        print("RadioControl function take either 'ON' or 'OFF' as variable")

    ser.close()
