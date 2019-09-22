
'''
Communications with the CX50, for manufacturing tests.

Documentation: http://wiki.gme.net.au/display/PDLC/CX50+Manufacturing+Software

Copyright (C) 2018 Standard Communications Pty Ltd (GME). All rights reserved.
'''

import sys
import time
import struct
import serial
import traceback
import cx50_manufacturing_test_log as mtl
import numpy as np


class CC1125(object):
    'CC1125 state representation'

    registers = (
        ('SYNC3', 0x0004),
        ('SYNC2', 0x0005),
        ('SYNC1', 0x0006),
        ('SYNC0', 0x0007),
        ('SYNC_CFG1', 0x0008),
        ('SYNC_CFG0', 0x0009),
        ('DEVIATION_M', 0x000A),
        ('MODCFG_DEV_E', 0x000B),
        ('DCFILT_CFG', 0x000C),
        ('PREAMBLE_CFG1', 0x000D),
        ('PREAMBLE_CFG0', 0x000E),
        ('FREQ_IF_CFG', 0x000F),
        ('IQIC', 0x0010),
        ('CHAN_BW', 0x0011),
        ('MDMCFG1', 0x0012),
        ('MDMCFG0', 0x0013),
        ('SYMBOL_RATE2', 0x0014),
        ('SYMBOL_RATE1', 0x0015),
        ('SYMBOL_RATE0', 0x0016),
        ('AGC_REF', 0x0017),
        ('AGC_CS_THR', 0x0018),
        ('AGC_GAIN_ADJUST', 0x0019),
        ('AGC_CFG3', 0x001A),
        ('AGC_CFG2', 0x001B),
        ('AGC_CFG1', 0x001C),
        ('AGC_CFG0', 0x001D),
        ('FIFO_CFG', 0x001E),
        ('DEV_ADDR', 0x001F),
        ('SETTLING_CFG', 0x0020),
        ('FS_CFG', 0x0021),
        ('WOR_CFG1', 0x0022),
        ('WOR_CFG0', 0x0023),
        ('WOR_EVENT0_MSB', 0x0024),
        ('WOR_EVENT0_LSB', 0x0025),
        ('PKT_CFG2', 0x0026),
        ('PKT_CFG1', 0x0027),
        ('PKT_CFG0', 0x0028),
        ('RFEND_CFG1', 0x0029),
        ('RFEND_CFG0', 0x002A),
        ('PA_CFG2', 0x002B),
        ('PA_CFG1', 0x002C),
        ('PA_CFG0', 0x002D),
        ('PKT_LEN', 0x002E),
        ('IF_MIX_CFG', 0x2F00),
        ('FREQOFF_CFG', 0x2F01),
        ('TOC_CFG', 0x2F02),
        ('MARC_SPARE', 0x2F03),
        ('ECG_CFG', 0x2F04),
        ('CFM_DATA_CFG', 0x2F05),
        ('EXT_CTRL', 0x2F06),
        ('RCCAL_FINE', 0x2F07),
        ('RCCAL_COARSE', 0x2F08),
        ('RCCAL_OFFSET', 0x2F09),
        ('FREQOFF1', 0x2F0A),
        ('FREQOFF0', 0x2F0B),
        ('FREQ2', 0x2F0C),
        ('FREQ1', 0x2F0D),
        ('FREQ0', 0x2F0E),
        ('IF_ADC2', 0x2F0F),
        ('IF_ADC1', 0x2F10),
        ('IF_ADC0', 0x2F11),
        ('FS_DIG1', 0x2F12),
        ('FS_DIG0', 0x2F13),
        ('FS_CAL3', 0x2F14),
        ('FS_CAL2', 0x2F15),
        ('FS_CAL1', 0x2F16),
        ('FS_CAL0', 0x2F17),
        ('FS_CHP', 0x2F18),
        ('FS_DIVTWO', 0x2F19),
        ('FS_DSM1', 0x2F1A),
        ('FS_DSM0', 0x2F1B),
        ('FS_DVC1', 0x2F1C),
        ('FS_DVC0', 0x2F1D),
        ('FS_LBI', 0x2F1E),
        ('FS_PFD', 0x2F1F),
        ('FS_PRE', 0x2F20),
        ('FS_REG_DIV_CML', 0x2F21),
        ('FS_SPARE', 0x2F22),
        ('FS_VCO4', 0x2F23),
        ('FS_VCO3', 0x2F24),
        ('FS_VCO2', 0x2F25),
        ('FS_VCO1', 0x2F26),
        ('FS_VCO0', 0x2F27),
        ('GBIAS6', 0x2F28),
        ('GBIAS5', 0x2F29),
        ('GBIAS4', 0x2F2A),
        ('GBIAS3', 0x2F2B),
        ('GBIAS2', 0x2F2C),
        ('GBIAS1', 0x2F2D),
        ('GBIAS0', 0x2F2E),
        ('IFAMP', 0x2F2F),
        ('LNA', 0x2F30),
        ('RXMIX', 0x2F31),
        ('XOSC5', 0x2F32),
        ('XOSC4', 0x2F33),
        ('XOSC3', 0x2F34),
        ('XOSC2', 0x2F35),
        ('XOSC1', 0x2F36),
        ('XOSC0', 0x2F37),
        ('ANALOG_SPARE', 0x2F38),
        ('PA_CFG3', 0x2F39),
        ('IRQ0M', 0x2F3F),
        ('IRQ0F', 0x2F40),
        ('WOR_TIME1', 0x2F64),
        ('WOR_TIME0', 0x2F65),
        ('WOR_CAPTURE1', 0x2F66),
        ('WOR_CAPTURE0', 0x2F67),
        ('BIST', 0x2F68),
        ('DCFILTOFFSET_I1', 0x2F69),
        ('DCFILTOFFSET_I0', 0x2F6A),
        ('DCFILTOFFSET_Q1', 0x2F6B),
        ('DCFILTOFFSET_Q0', 0x2F6C),
        ('IQIE_I1', 0x2F6D),
        ('IQIE_I0', 0x2F6E),
        ('IQIE_Q1', 0x2F6F),
        ('IQIE_Q0', 0x2F70),
        ('RSSI1', 0x2F71),
        ('RSSI0', 0x2F72),
        ('MARCSTATE', 0x2F73),
        ('LQI_VAL', 0x2F74),
        ('PQT_SYNC_ERR', 0x2F75),
        ('DEM_STATUS', 0x2F76),
        ('FREQOFF_EST1', 0x2F77),
        ('FREQOFF_EST0', 0x2F78),
        ('AGC_GAIN3', 0x2F79),
        ('AGC_GAIN2', 0x2F7A),
        ('AGC_GAIN1', 0x2F7B),
        ('AGC_GAIN0', 0x2F7C),
        ('CFM_RX_DATA_OUT', 0x2F7D),
        ('CFM_TX_DATA_IN', 0x2F7E),
        ('ASK_SOFT_RX_DATA', 0x2F7F),
        ('RNDGEN', 0x2F80),
        ('MAGN2', 0x2F81),
        ('MAGN1', 0x2F82),
        ('MAGN0', 0x2F83),
        ('ANG1', 0x2F84),
        ('ANG0', 0x2F85),
        ('CHFILT_I2', 0x2F86),
        ('CHFILT_I1', 0x2F87),
        ('CHFILT_I0', 0x2F88),
        ('CHFILT_Q2', 0x2F89),
        ('CHFILT_Q1', 0x2F8A),
        ('CHFILT_Q0', 0x2F8B),
        ('GPIO_STATUS', 0x2F8C),
        ('FSCAL_CTRL', 0x2F8D),
        ('PHASE_ADJUST', 0x2F8E),
        ('PARTNUMBER', 0x2F8F),
        ('PARTVERSION', 0x2F90),
        ('SERIAL_STATUS', 0x2F91),
        ('MODEM_STATUS1', 0x2F92),
        ('MODEM_STATUS0', 0x2F93),
        ('MARC_STATUS1', 0x2F94),
        ('MARC_STATUS0', 0x2F95),
        ('PA_IFAMP_TEST', 0x2F96),
        ('FSRF_TEST', 0x2F97),
        ('PRE_TEST', 0x2F98),
        ('PRE_OVR', 0x2F99),
        ('ADC_TEST', 0x2F9A),
        ('DVC_TEST', 0x2F9B),
        ('ATEST', 0x2F9C),
        ('ATEST_LVDS', 0x2F9D),
        ('ATEST_MODE', 0x2F9E),
        ('XOSC_TEST1', 0x2F9F),
        ('XOSC_TEST0', 0x2FA0),
        ('RXFIRST', 0x2FD2),
        ('TXFIRST', 0x2FD3),
        ('RXLAST', 0x2FD4),
        ('TXLAST', 0x2FD5),
        ('NUM_TXBYTES', 0x2FD6),
        ('NUM_RXBYTES', 0x2FD7),
        ('FIFO_NUM_TXBYTES', 0x2FD8),
        ('FIFO_NUM_RXBYTES', 0x2FD9),
        ('SINGLE_TXFIFO', 0x003F),
        ('BURST_TXFIFO', 0x007F),
        ('SINGLE_RXFIFO', 0x00BF),
        ('BURST_RXFIFO', 0x00FF),
    )

    MODE_UNDEFINED = -1
    MODE_IDLE = 0
    MODE_RX = 1
    MODE_TX = 2

    def __init__(self):
        self.tx_power = 0
        self.freq = 0
        self.mode = self.MODE_UNDEFINED


class CommsError(Exception):
    pass


class CommsTimeout(Exception):
    pass


class GME2(object):
    'Single-threaded GME2 protocol implementation'

    # GME2 special packet characters
    ESC_CHAR = 0x7D # Escape char
    STX_CHAR =  0x7E # Start of packet
    ETX_CHAR = 0x7F # End of packet

    # GME2 commands
    CMD_READ_VERSION = 0x00  #  Read firmware version.
    CMD_READ_EE_BYTE = 0x01  #  Read EEPROM byte.
    CMD_READ_EE_WORD = 0x02  #  Read EEPROM word.
    CMD_READ_EE_BLOCK = 0x03  #  Read EEPROM block (16-bytes).
    CMD_GET_CHANNEL = 0x04  #  Read current channel.
    CMD_READ_STATUS = 0x05  #  Read status flags.
    CMD_READ_OPFLAGS = 0x06  #  Read operation flags. (not used)
    CMD_READ_EE_ARRAY = 0x0B  #  Read array of bytes from EEPROM. (not used)
    CMD_READ_ID = 0x0C  #  Read Radio ID.
    CMD_READ_MAC_ADDR = 0x0D  #  Read Radio MAC Addr.
    CMD_READ_SRP_STATUS = 0x0E  #  Read SRP Enabled
    CMD_SET_SRP_ENABLED = 0x0F  #  only connect via SRP
    CMD_SET_DEST_MODE = 0x10  #  Change RH data route from radio to MSM
    CMD_MSM_CTL = 0x11  #  Issue control action related to MSM
    CMD_MSM_DATA = 0x12  #  Data exchanged between PC Programmer and MSM
    CMD_READ_FEATURE_FLAGS = 0x13   #  Read back from radio the feature flag bits.
    CMD_WRITE_EE_BYTE = 0x21   #  Write EEPROM byte.
    CMD_WRITE_EE_WORD = 0x22   #  Write EEPROM word.
    CMD_WRITE_EE_BLOCK = 0x23   #  Write EEPROM block (16-bytes).
    CMD_SET_CHANNEL = 0x24   #  Set current channel.
    CMD_WRITE_STATUS = 0x25   #  Write status flags
    CMD_SET_PTT = 0x26   #  PTT on/off.
    CMD_SET_SQUELCH = 0x27   #  Set squelch level SQL[1-9].
    CMD_SENDKEY = 0x28   #  Send remote keycode. (not used)
    CMD_CH_UP = 0x29   #  Change up to next available channel.
    CMD_CH_DN = 0x2A   #  Change down to next available channel.
    CMD_WRITE_EE_ARRAY = 0x2B   #  Write array of bytes to EEPROM. (not used)
    CMD_WRITE_ID = 0x2C   #  Write Radio ID
    CMD_SET_VOLUME = 0x2F   #  Set volume [0-31].
    CMD_TEST = 0x34   #  Test command. Prefix for test sub-command.
    CMD_UIC = 0x35   #  Commands related to UIC
    CMD_SRP_AUTH = 0x36   #  SRP6a based Authentication.
    CMD_SEND_LICENSE_KEY = 0x37   #  Updates the feature flags in the Radio.
    CMD_SETMODE = 0x38   #  Set TEST mode.
    CMD_PROG_MODE = 0x39   #  Set PROG mode.
    CMD_READ_RADIO_TYPE = 0x3A   #  Read the radio type.
    CMD_SET_BAUD_RATE = 0x3B
    CMD_MSM = 0x3C   #  MSM command.
    CMD_REMOTE = 0x3D   #  Wireless remote command.
    CMD_ENTER_BOOTMODE = 0x3E   #  Software reset to bootloader.
    CMD_RESET = 0x3F   #  Software reset.

    # MSM Control subcommands (Related to 'CMD_MSM_CTL' (0x11))
    CMD_MSM_CTL_OPEN_CHN = 0x01   #  Open communication channel between PC and MSM (including GME2 <-> SLIP conversion)
    CMD_MSM_CTL_CLOSE_CHN = 0x02   #  Close communication channel between PC and MSM opened by _OPEN_PIPE.
    CMD_MSM_CTL_RESET_HW = 0x03   #  Hardware reset MSM
    CMD_MSM_CTL_RESET_SW = 0x04   #  Issue software reset to MSM
    CMD_MSM_CTL_DBG = 0x05   #  Debug the protocol converter
    CMD_MSM_CTL_SKIP_INIT = 0x06   #  Skip MSM initialisation
    CMD_MSM_CTL_TRA_INIT = 0x07   #  Initialize data transformation defined by TRA field
    CMD_MSM_CTL_RD_STA = 0x08   #  Read MSM presence flag

    # MSM test commands (currently used for CM-60 modulation accuracy test)
    CMD_MSM_TEST = 0x00
    CMD_MSM_TEST_SETUP = 0x00 # Set the MSM frequency and bandwidth
    CMD_MSM_TEST_MODE = 0x01 # Set the MSM test mode

    # MSM test mode (used in the command above)
    MSM_TEST_MODE_DEACTIVATE = 0 # Deactivate All, deactivate all test mode functions
    MSM_TEST_MODE_TONE = 1 # Tone, transmit the standard tone test pattern (IMBE 1011 Hz pattern)
    MSM_TEST_MODE_SILENCE = 2 # Silence, transmit the standard silence test pattern (IMBE Silence pattern)
    MSM_TEST_MODE_TX_TEST = 3 # TX Test, transmit the standard transmitter test pattern ([S-42] 1.3.4.3)
    MSM_TEST_MODE_TX_SYMBOL_RATE = 4 # TX Symbol Rate, transmit the standard transmitter symbol rate pattern ([S-42] 1.3.4.4)
    MSM_TEST_MODE_TX_LOW_DEVIATION = 5 # TX Low Deviation, transmit the standard transmitter low deviation pattern ([S-42] 1.3.4.5)
    MSM_TEST_MODE_C4FM_MOD_FIDELITY = 6 # C4FM Mod Fidelity, transmit the standard transmitter C4FM modulation fidelity pattern ([S-42] 1.3.4.6)
    MSM_TEST_MODE_RX_CALC_SCALE_FACTORS = 11 # RX Calc Scale Factors, calculate the RX and RSSI scaling factors
    MSM_TEST_MODE_RX_CON_CALC_SCALE_FACTORS = 12 # RX Con Calc Scale Factors, continuously calculate the RX and RSSI scaling factors.

    # Test mode subcommands
    CMD_READ_MEM = 0x01   #  Read RAM location.
    CMD_READ_MEM_INT = 0x02   #  Read RAM location (16 bit)
    CMD_SET_BEEP_VOL = 0x06   #  Set beep tone volume
    CMD_READ_DAC = 0x07   #  Read DAC hardware register. (not used)
    CMD_READ_ADC = 0x08   #  Read A/D channel.
    CMD_READ_GPIO = 0x09 # Read GPIO line directly
    CMD_WRITE_GPIO = 0x0A # Write GPIO line directly
    CMD_SET_PWRMG = 0x0B # Enable/disable power management
    CMD_GET_FREQUENCY = 0x0C  # Get Tx/Rx frequency
    CMD_WRITE_PLLMUX = 0x0D # Write PLL mux out
    CMD_READ_PLL_UNLOCK = 0x0E # Read "PLL unlock" pin status
    CMD_WRITE_PLL_FREQ = 0x0F # Write PLL frequency directly (bypass TX/RX logic, etc.)
    CMD_SET_LCD = 0x10 # Set LCD on/off
    DSP_TEST_CAPTURE_START = 0x11 # Capture the DSP test buffer
    DSP_TEST_CAPTURE_READ = 0x12 # Read out the DSP test buffer
    CMD_READ_PLL_FREQ = 0x13 # Read PLL frequency directly
    CMD_READ_RCM_STATE = 0x14 # Read the current RCM state
    DSP_TEST_CONFIG = 0x15 # Ad-hoc DSP test config, for development only
    CMD_CC1125 = 0x16 # CC1125 test subcommand (see below)
    QRY_CTCSS_STATUS = 0x17 # Query the status of the primary CTCSS decoder
    CMD_GET_TRACE = 0x18 # Get debuging trace
    CMD_WRITE_MEM = 0x81   #  Write to RAM location.
    SET_SELCALL_TXTONE = 0x82   #  Set selcall TX test tone.
    SET_CTCSS_TXTONE = 0x83   #  Set CTCSS TX test tone.
    WRITE_LCD = 0x84   #  Write directly to LCD.
    WRITE_PLL = 0x85   #  Write directly to PLL registers.
    SET_SQUELCH = 0x86   #  Squelch On/Off.
    CMD_WRITE_DAC = 0x87   #  Write to DAC hardware register.
    CMD_SET_FREQUENCY = 0x88   #  Set Tx/Rx frequency
    AUTOALIGN = 0x89   #  Frontend autoalignment sweep.
    CMD_BEEPTONE = 0x8A   #  Beeptone test.
    READ_AUDIO = 0x8B   #  Read received audio samples.
    READ_FREQ_ERROR = 0x8C   #  Read receive frequency error.
    CMD_TEST_OPTIONS = 0x8D   #  Temporary override various system flags
    RGB_BKLIGHT = 0x8E # Set RGB Backlight values.
    RX_OFFSET_ALIGN = 0x8F # Self calibrate receive frequency offset.
    CMD_READ_OUT_AUDIO = 0x90 # Read filtered output audio.
    CMD_SET_TX_TONE = 0x91
    CMD_SET_FREQ_x50Hz = 0x92
    CMD_SET_TX_PWR_LEVEL = 0x93 # Set Tx power level: (1, 5, 10, 25W (0x00, 0x01, 0x02, 0x03) respectively)
    CMD_SET_UART_RELAY = 0x94 # Set pass-through mode between two UARTs
    CMD_SELECT_PA = 0x95 # Select PA: (0 = low power, 1 = high power)
    CMD_READ_BPF_ALIGN = 0x9F
    CMD_LCD_TEST_PATTERN = 0xB0
    CMD_GET_KEY = 0xB1
    CMD_GET_BIST_FLAGS = 0xB2  # Get built-in self test result flags

    # Key codes
    UP_KEY_H = 0x00  # Rotary knob UP
    DN_KEY_H = 0x01  # Rotary knob DN
    F1_KEY_H = 0x02
    F2_KEY_H = 0x03
    F3_KEY_H = 0x04
    F4_KEY_H = 0x05
    PTT_KEY_H = 0x06
    A_KEY_H = 0x08
    B_KEY_H = 0x09
    EMERG_KEY_H = 0x0A  # Red emergency key
    CH_UP_KEY_H = 0x0C  # P/B UP
    CH_DN_KEY_H = 0x0D  # P/B DN
    QB1_KEY_H = 0x20  # CX50 QB1
    QB2_KEY_H = 0x21  # CX50 QB2
    QB3_KEY_H = 0x22  # CX50 QB3
    TILT_KEY_H = 0x30  # Fake key for tilt switch.
    NO_KEY_H = 0xFF

    # Key operations
    KEY_PRESS = 0x03
    KEY_HELD = 0x05
    KEY_RELEASE = 0x01
    KEY_INVALID = 0x00

    # UIC remote sub-commands.
    UIC_KEY_IND = 0x4B
    UIC_DISP_SET_SCREEN_REQ = 0x40
    UIC_DISP_N_LIST_ITEMS_REQ = 0x45
    UIC_DISP_GET_LIST_ITEMS_IND = 0x46
    UIC_DISP_LIST_ITEMS_REQ = 0x47

    # Wireless remote sub-commands.
    W_CTRL = 0x00   #  Control command
    W_KEY = 0x01   #  Wireless key.
    W_VERSION = 0x02   #  Wireless firmware version.
    W_ADDR_RCV = 0x00   #  Address receive.
    W_ADDR_WR = 0x01   #  Address write.
    W_ADDR_RD = 0x02   #  Address read.

    # MSM sub-commands
    CMD_MSM_TEST = 0x00   #  Test command subset
    CMD_MSM_TEST_SETUP = 0x00   #  Set the MSM frequency and bandwidth
    CMD_MSM_TEST_MODE = 0x01   #  Set the MSM test mode
    CMD_MSM_TEST_RX_STATS = 0x02   #  Rx stats packet received from MSM

    # SRP Authentication sub-commands
    CMD_SRP_AUTH_START = 0x00   # Start SRP authentication
    CMD_SRP_AUTH_DATA = 0x01   # SRP authentication data
    CMD_SRP_AUTH_SUCCESS = 0x02 # SRP authentication success
    CMD_SRP_AUTH_STOP = 0x03 # Stop SRP authentication
    CMD_SRP_ENCR_DATA = 0x04 # Encrypted SRP data
    CMD_SRP_AUTH_SALT = 0x05 # SRP authentication data

    # EEPROM calibration addresses
    EE_PRODUCT_ID = 0  # Product ID code (16-bit).
    EE_PRODUCT_ID_1 = 1

    # RX tracking filter alignment points -------
    EE_T1_F0 = 912 # 403 MHz Alignment
    EE_T2_F0 = 913
    EE_T3_F0 = 914
    EE_T4_F0 = 915

    EE_T1_F1 = 2  # 450 MHz Alignment
    EE_T2_F1 = 3
    EE_T3_F1 = 4
    EE_T4_F1 = 5

    EE_T1_F2 = 6  # 477 MHz Alignment
    EE_T2_F2 = 7
    EE_T3_F2 = 8
    EE_T4_F2 = 9

    EE_T1_F3 = 10 # 520 MHz Alignment
    EE_T2_F3 = 11
    EE_T3_F3 = 12
    EE_T4_F3 = 13

    EE_T1_F4 = 908 # 520 MHz Alignment
    EE_T2_F4 = 909
    EE_T3_F4 = 910
    EE_T4_F4 = 911

    # Noisevolts calibration
    EE_NV_NB_SCALE = 942  # Narrowband scaling factor for ADC
    EE_NV_WB_SCALE = 944  # Wideband scaling factor for ADC
    EE_NV_RESERVED = 946  # 4 bytes reserved until NV work on CM60 completed

    # -------------------------------------------

    EE_MOD = 14 # Mod offset (16)
    EE_SQL_K = 15 # Squelch threshold calibration. (-1) [FF]
    EE_BAL = 16 # Balance   (12)
    EE_AFC_LIMIT = 17 # AFC tuning range.(192)
    EE_PLL_FREF = 20 # Phase Lock Loop Ref. Freq. offset (120)
    EE_XTAL_TEMPCO = 21 # Temperature compensation linear term. (255)
    EE_S1_CAL = 22 # S-Meter S1 Calibration Point
    EE_S10_CAL = 23 # S-Meter S10 Calibration Point
    EE_XTAL_TEMPCO_B = 24 # Temperature compensation cubic term. (255)
    EE_CONFIG = 25 # Configuration flags. (0xFF)
    EE_RX_DEMOD_ADJ = 26 # Rx DC offset adjust by 8 bit tuning DAC for analog channels (1 byte)
    EE_REF_RX_OFFSET = 27 # Calibration offset for FREF when receiving (+/-15)
    EE_PCB_REV = 28 # PCB Revision Code (16-bit).
    EE_PCB_REV_1 = 29
    EE_STATUS = 30 # Status flags (0)

    EE_SQL_LEVEL = 32 # SQL level 1-9
    EE_TX_TIMEOUT = 33

    EE_MIC_GAIN_AUX = 106 # AUX port line gain
    EE_MIC_GAIN_WIRELESS = 107 # Remote PTT MIC gain. +24/-12dB, 1.0dB steps.         [00] 0dB
    EE_CTCSS_WB_DEVIATION = 108 # CTCSS WB TX deviation. 0-80 -> 0..800Hz
    EE_DCS_WB_DEVIATION = 109 # DCS WB TX deviation. 0-80 -> 0..800Hz
    EE_MIC_GAIN_FRONT = 110 # Front PTT MIC gain. +24/-12dB, 1.0dB steps.          [00] 0dB
    EE_MIC_GAIN_REAR = 111 # Rear PTT MIC gain. +24/-12dB, 1.0dB steps.           [00] 0dB
    EE_CTCSS_NB_DEVIATION = 112 # CTCSS NB TX deviation. 0-80 -> 0..800Hz
    EE_DCS_NB_DEVIATION = 113 # DCS NB TX deviation. 0-80 -> 0..800Hz
    EE_CTCSS2_TONE = 114
    EE_SPK_GAIN_SW = 115 # SW SPEAKER gain. +24/-12dB, 1.0dB steps.             [00] 0dB

    # TX/PA calibration
    EE_MOD_CAL_1 = 980
    EE_MOD_CAL_2 = 981
    EE_MOD_CAL_3 = 982
    EE_MOD_CAL_4 = 983
    EE_TX_AUX_POWER_1 = 984
    EE_TX_AUX_POWER_2 = 985
    EE_TX_AUX_POWER_3 = 986
    EE_TX_AUX_POWER_4 = 987
    EE_TX_POWER_1 = 988
    EE_TX_POWER_2 = 989
    EE_TX_POWER_3 = 990
    EE_TX_POWER_4 = 991
    EE_PWR_SCALE_1W = 992

    # Test options (CMD_TEST_OPTIONS) flags
    TEST_OPTIONS_AFC = 0x01 #  b0: Turn AFC on/off.
    TEST_OPTIONS_WB_DEVIATION = 0x02 # b1: Wideband/Narrowband deviation select.
    TEST_OPTIONS_WB_FILTER = 0x04 # b2: Wideband/Narrowband IF filter select.

    # GPIO IDs for CMD_READ_GPIO and CMD_WRITE_GPIO
    GPIO_VCO_OFFSET = 0x01
    GPIO_BT_P2_0 = 0x02
    GPIO_FE_PA_OUT1 = 0x03
    GPIO_FE_PA_OUT2 = 0x04
    GPIO_FE_SW1_V2 = 0x05
    GPIO_FE_SW1_V1 = 0x06
    GPIO_LED_RX_ON = 0x07
    GPIO_LED_TX_ON = 0x08
    GPIO_TRX_RFA_ON = 0x09
    GPIO_TRX_RX_ON = 0x0A
    GPIO_TRX_TX_ON = 0x0B
    GPIO_TRX_PLL_ON = 0x0C
    GPIO_WB_ON = 0x0D
    GPIO_TP106 = 0x0E
    GPIO_AUDIO_DISABLE = 0x0F
    GPIO_ids = (
        GPIO_VCO_OFFSET,
        GPIO_BT_P2_0,
        GPIO_FE_PA_OUT1,
        GPIO_FE_PA_OUT2,
        GPIO_FE_SW1_V2,
        GPIO_FE_SW1_V1,
        GPIO_LED_RX_ON,
        GPIO_LED_TX_ON,
        GPIO_TRX_RFA_ON,
        GPIO_TRX_RX_ON,
        GPIO_TRX_TX_ON,
        GPIO_TRX_PLL_ON,
        GPIO_WB_ON,
        GPIO_TP106,
        GPIO_AUDIO_DISABLE)

    # A/D Channels (CMD_READ_ADC)
    AD_BATT_MON = 0
    AD_TEMPERATURE = 1
    AD_MIC_AUDIO = 3
    AD_RX_AUDIO = 4
    AD_KEYS_TOP = 8
    AD_NOISE_MON = 9
    AD_TRX_RX_RSSI = 10
    AD_KEYS_BOTTOM = 15
    AD_KEYS_QB = 14
    ADC_channels = (
        AD_BATT_MON,
        AD_TEMPERATURE,
        AD_MIC_AUDIO,
        AD_RX_AUDIO,
        AD_KEYS_TOP,
        AD_NOISE_MON,
        AD_TRX_RX_RSSI,
        AD_KEYS_BOTTOM,
        AD_KEYS_QB)

    # DAC channels (CMD_READ/WRITE_DAC)
    HW_TUNING_FREF_PORT = 1
    HW_TUNING_TP507 = 2
    HW_TUNING_TX_POWER_ADJ_PORT = 3
    HW_TUNING_RX_BFP_TUNE1_PORT = 4
    HW_TUNING_RX_BFP_TUNE2_PORT = 5
    HW_TUNING_RX_BFP_TUNE3_PORT = 6
    HW_TUNING_TP502 = 7
    HW_TUNING_BK_LED = 8
    DAC_channels = (
        HW_TUNING_FREF_PORT,
        HW_TUNING_TP507,
        HW_TUNING_TX_POWER_ADJ_PORT,
        HW_TUNING_RX_BFP_TUNE1_PORT,
        HW_TUNING_RX_BFP_TUNE2_PORT,
        HW_TUNING_RX_BFP_TUNE3_PORT,
        HW_TUNING_TP502,
        HW_TUNING_BK_LED)

    # PLL MUX OUT values
    PLL_MUX_HIGH_Z = 0
    PLL_MUX_DLD = 1
    PLL_MUX_N_DIV = 2
    PLL_MUX_AVDD = 3
    PLL_MUX_R_DIV = 4
    PLL_MUX_ALD = 5
    PLL_MUX_SDO = 6
    PLL_MUX_DGND = 7
    PLL_MUX_ids = (
        PLL_MUX_HIGH_Z,
        PLL_MUX_DLD,
        PLL_MUX_N_DIV,
        PLL_MUX_AVDD,
        PLL_MUX_R_DIV,
        PLL_MUX_ALD,
        PLL_MUX_SDO,
        PLL_MUX_DGND)

    # DSP test capture source IDs
    DTC_RX_ADC = 0  # RX, ADC (discriminator) samples
    DTC_RX_DEC = 1  # RX, after input decimation
    DTC_RX_DCR = 2  # RX, after DC removal
    DTC_RX_HPF = 3  # RX, after subtone-removal high pass filter
    DTC_RX_DAC = 4  # RX, DAC samples (after interpolation)
    DTC_RX_SLC = 5  # RX, selcall tone decoder output
    DTC_TX_ADC = 6  # TX, ADC (mike) samples
    DTC_TX_DEC = 7  # TX, after input decimation
    DTC_TX_DCR = 8  # TX, after DC removal and gain
    DTC_TX_HII = 9  # TX, Hilbert, I
    DTC_TX_HIQ = 10 # TX, Hilbert, Q
    DTC_TX_MAG = 11 # TX, magnitude estimation
    DTC_TX_CLP = 12 # TX, after clipping
    DTC_TX_SIG = 13 # TX, after adding tones and signalling
    DTC_TX_DAC = 14 # TX, DAC (modulation) samples
    DTC_ids = (
        DTC_RX_ADC,
        DTC_RX_DEC,
        DTC_RX_DCR,
        DTC_RX_HPF,
        DTC_RX_DAC,
        DTC_RX_SLC,
        DTC_TX_ADC,
        DTC_TX_DEC,
        DTC_TX_DCR,
        DTC_TX_HII,
        DTC_TX_HIQ,
        DTC_TX_MAG,
        DTC_TX_CLP,
        DTC_TX_SIG,
        DTC_TX_DAC)

    # RCM states
    RCM_ST_INIT = 0
    RCM_ST_RECEIVE = 1
    RCM_ST_BUSY_CHECK = 2
    RCM_ST_TX_SETUP = 3
    RCM_ST_TRANSMITTING = 4
    RCM_ST_TX_FINISH = 5
    RCM_ST_POWER_DOWN = 6
    RCM_ST_PTT_FAILURE = 7

    # Test subcommands for CMD_CC1125 (see above)
    CMD_CC1125_RESET = 0x01       # Reset the chip (normally done at initialisation)
    CMD_CC1125_SLEEP = 0x02       # Switch to sleep mode
    CMD_CC1125_IDLE = 0x03        # Switch to idle mode
    CMD_CC1125_INIT_RX = 0x04     # Initialize RX for RSSI measurement (it's not a mode through.)
    CMD_CC1125_TX = 0x05          # Switch to transmit (CW) mode.
    CMD_CC1125_GET_STATUS = 0x06  # Read CC1125 status
    CMD_CC1125_READ = 0x07        # Read CC1125 register. Following 2 bytes: address-hi, address-lo.
    CMD_CC1125_WRITE = 0x08       # Write CC1125 register. Following 3 bytes: address-hi, address-lo, value.
    CMD_CC1125_FREQ = 0x09        # Set Tx or Rx frequency. Following 4 bytes: frequency register values (sorry, not in Hz)
    CMD_CC1125_TX_PWR = 0x0A      # Set the transmit output power. Following 2 bytes: power-hi, power-lo.
    CMD_CC1125_GET_RSSI = 0x0B    # Report the latest RSSI.

    # Misc constants
    DEFAULT_RADIO_BAUD_RATE = 9600
    CHECKSUM_XOR_MASK = 0xFF
    PKT_CMD_NACK_BIT = 0x40
    PLL_STEP = 0.0125 # MHz

    # Calibration EEPROM address tables
    N_CAL_DACS = 3 # RX calibration DAC values (varactors) on each frequency
    CAL_FREQ_0 = 403
    CAL_FREQ_1 = 450
    CAL_FREQ_2 = 477
    CAL_FREQ_3 = 504
    CAL_FREQ_4 = 520
    rx_cal_address = {}
    rx_cal_address[CAL_FREQ_0] = (EE_T1_F0, EE_T2_F0, EE_T3_F0)
    rx_cal_address[CAL_FREQ_1] = (EE_T1_F1, EE_T2_F1, EE_T3_F1)
    rx_cal_address[CAL_FREQ_2] = (EE_T1_F2, EE_T2_F2, EE_T3_F2)
    rx_cal_address[CAL_FREQ_3] = (EE_T1_F3, EE_T2_F3, EE_T3_F3)
    rx_cal_address[CAL_FREQ_4] = (EE_T1_F4, EE_T2_F4, EE_T3_F4)

    tx_mod_address = {}
    tx_mod_address[CAL_FREQ_1] = EE_MOD_CAL_1
    tx_mod_address[CAL_FREQ_2] = EE_MOD_CAL_2
    tx_mod_address[CAL_FREQ_3] = EE_MOD_CAL_3
    tx_mod_address[CAL_FREQ_4] = EE_MOD_CAL_4

    tx_aux_power_cal_address = {}
    tx_aux_power_cal_address[CAL_FREQ_1] = EE_TX_AUX_POWER_1
    tx_aux_power_cal_address[CAL_FREQ_2] = EE_TX_AUX_POWER_2
    tx_aux_power_cal_address[CAL_FREQ_3] = EE_TX_AUX_POWER_3
    tx_aux_power_cal_address[CAL_FREQ_4] = EE_TX_AUX_POWER_4

    tx_power_cal_address = {}
    tx_power_cal_address[CAL_FREQ_1] = EE_TX_POWER_1
    tx_power_cal_address[CAL_FREQ_2] = EE_TX_POWER_2
    tx_power_cal_address[CAL_FREQ_3] = EE_TX_POWER_3
    tx_power_cal_address[CAL_FREQ_4] = EE_TX_POWER_4


    def __init__(self, port):
        self.log = mtl.get_logger('cx50_control.gme2')
        self.port = port
        self.response_timeout = 0.5


    def receive_packet(self, timeout):
        'Wait for, and receive GME2 packet'
        # read(1) below has to be non-blocking so that we can exit this thread
        # Let's use the packet length itself as the "state variable"
        packet = bytearray()
        while True:
            if time.clock() > timeout:
                raise CommsTimeout()

            # Max 10ms blocking timeout per byte
            c = self.port.read(1)
            if not c:
                continue

            b = ord(c)

            # Byte 0: STX
            if len(packet) == 0 and b == GME2.STX_CHAR:
                packet.append(b)
                continue

            # Byte 1 must be == 2
            if len(packet) == 1:
                if b == 2:
                    packet.append(b)
                    continue
                else:
                    # Log and discard it
                    self.log.debug(mtl.hex_message('<--', packet))
                    raise CommsError('Received invalid packet type.')

            # Byte 2: payload length
            if len(packet) == 2:
                packet.append(b)
                payload_size = b
                escape = False
                checksum = 0
                continue

            # Next few bytes: the payload + checksum. Un-escape the payload and do the checksum.
            if len(packet) > 2 and len(packet) < payload_size + 3:
                if b == GME2.ESC_CHAR:
                    escape = True
                else:
                    if escape:
                        b ^= 0x20
                    escape = False
                    packet.append(b)
                    checksum ^= b
                continue

            # Next byte after payload: Checksum
            if len(packet) > 2 and len(packet) == payload_size + 3:
                checksum ^= self.CHECKSUM_XOR_MASK
                packet.append(b)
                if checksum == b:
                    continue
                else:
                    self.log.debug(mtl.hex_message('<--', packet))
                    raise CommsError('Failed GME2 comms checksum (received: %02X, calculated: %02X)' % (b, checksum))

            # Last byte should be the ETX
            if len(packet) >= 2 and len(packet) == payload_size + 4:
                if b == GME2.ETX_CHAR:
                    packet.append(b)
                    # If this is an echo (TX7200 cable does this), discard it and wait for the actual response
                    if self.echo_data and self.echo_data == packet:
                        self.log.debug('<-- Echo:')
                        self.log.debug(mtl.hex_message('<--', packet))
                        self.echo_data = None
                        packet = bytearray()
                        continue
                    else:
                        self.log.debug(mtl.hex_message('<--', packet))
                        payload = packet[3:-2]
                        packet = bytearray()
                        return payload
                else:
                    raise CommsError('<-- Malformed packet; missing ETX.')


    def send(self, *payload):
        'Package up and and send a GME2 packet'
        # Header
        tx_array = bytearray((GME2.STX_CHAR, 2, len(payload)))

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
            if bb in (GME2.ESC_CHAR, GME2.STX_CHAR, GME2.ETX_CHAR):
                tx_array.append(GME2.ESC_CHAR)
                tx_array.append(bb ^ 0x20)
            else:
                tx_array.append(bb)

        # Trailer
        tx_array.append(chksum ^ self.CHECKSUM_XOR_MASK)
        tx_array.append(GME2.ETX_CHAR)

        # TX7200 (the USB dongle, actually) echoes everything back
        self.echo_data = tx_array

        # Log & transmit
        self.log.debug(mtl.hex_message('-->', tx_array))
        self.port.write(tx_array)


    def request(self, *data):
        'Send a request to the radio, wait for a response and return the response.'
        self.send(*data)
        timeout = time.clock() + self.response_timeout
        while time.clock() < timeout:
            resp = self.receive_packet(timeout)
            # Wait for a response to the last command
            # (CM60, for example, sends out an unsolicited ping once per second)
            if resp:
                if resp[0] & GME2.PKT_CMD_NACK_BIT:
                    raise CommsError('Radio sent a NACK.')
                elif resp[0] == data[0]:
                    return resp
            time.sleep(0.01)
        raise CommsTimeout()


class CX50(object):
    def __init__(self, com_port):
        self.log = mtl.get_logger('cx50_control.cx50')
        self.port = com_port
        self.gme2 = GME2(self.port)

    def check_connection(self):
        # Which baud rate is the radio using?
        current_baud_rate = None
        available_baud_rates = 9600, 19200, 38400, 57600, 115200, 230400, 1000000
        for baud_rate in available_baud_rates:
            self.port.baudrate = baud_rate
            self.port.reset_input_buffer()
            try:
                resp = self.gme2.request(GME2.CMD_READ_VERSION)
            except CommsTimeout:
                if baud_rate == available_baud_rates[-1]:
                    # We've now tried all the rates and still no response
                    raise CommsTimeout('No response from the radio.')
                else:
                    time.sleep(0.1)
                    continue
            current_baud_rate = baud_rate
            if len(resp) > 3:
                self.firmware_version_string = '%d.%02d%c' % (resp[1], resp[2], resp[3])
            else:
                self.firmware_version_string = '%d.%02d' % (resp[1], resp[2])
            break
        # Change to desired baud rate
        fast_baud_rate = 115200
        if current_baud_rate != fast_baud_rate:
            # Let's try switching to the faster rate now
            self.gme2.request(
                GME2.CMD_SET_BAUD_RATE,
                (fast_baud_rate >> 16) & 0xFF,
                (fast_baud_rate >> 8) & 0xFF,
                fast_baud_rate & 0xFF)
        # Check comms at the new speed
        self.port.baudrate = fast_baud_rate
        self.gme2.request(GME2.CMD_PROG_MODE)
        # Always keep power management disabled while connected to the alignment tool
        self.gme2.request(GME2.CMD_TEST, GME2.CMD_SET_PWRMG, 0)
        # Always enable test mode while connected to the test script
        self.gme2.request(GME2.CMD_SETMODE, 1)

    def read_adc(self, channel):
        if channel not in GME2.ADC_channels:
            raise CommsError('Coding error - incorrect ADC channel.')
        resp = self.gme2.request(GME2.CMD_TEST, GME2.CMD_READ_ADC, channel)
        return int(resp[3])

    def read_eeprom_byte(self, address):
        resp = self.gme2.request(GME2.CMD_READ_EE_BYTE, (address >> 8) & 0xFF, address & 0xFF)
        return resp[3]

    def write_eeprom_byte(self, address, value):
        return self.gme2.request(GME2.CMD_WRITE_EE_BYTE, (address >> 8) & 0xFF, address & 0xFF, value)

    def read_eeprom_word(self, address):
        lb = self.read_eeprom_byte(address)
        hb = self.read_eeprom_byte(address + 1)
        return (hb << 8) | lb

    def write_eeprom_word(self, address, value):
        lb = value & 0xFF
        hb = (value >> 8) & 0xFF
        self.write_eeprom_byte(address, lb)
        self.write_eeprom_byte(address + 1, hb)

    def read_eeprom_array(self, address, size):
        resp = self.gme2.request(GME2.CMD_READ_EE_ARRAY , (address >> 8) & 0xFF, address & 0xFF, size & 0xFF)
        return resp[3:]

    def adc_batt_volts(self):
        adc_val = self.read_adc(GME2.AD_BATT_MON)
        batt_volts = adc_val * (3.3 / 255) * (1 + 100e3 / 56e3)
        return batt_volts

    def rssi_adc(self):
        return self.read_adc(GME2.AD_TRX_RX_RSSI)

    def noise_adc(self):
        return self.read_adc(GME2.AD_NOISE_MON)

    def mike_adc(self):
        return self.read_adc(GME2.AD_MIC_AUDIO)

    def rx_adc(self):
        return self.read_adc(GME2.AD_RX_AUDIO)

    def adc_keys_top(self):
        return self.read_adc(GME2.AD_KEYS_TOP)

    def adc_keys_bottom(self):
        return self.read_adc(GME2.AD_KEYS_BOTTOM)

    def adc_keys_qb(self):
        return self.read_adc(GME2.AD_KEYS_QB)

    def adc_temperature(self):
        temp_adc = self.read_adc(GME2.AD_TEMPERATURE)
        # ADC value to temperature celsius
        Rth = 10e3 * temp_adc / (255 - temp_adc)
        R25 = 10e3
        B = 3980.0
        Kelvin_offset = 273.15
        T25 = Kelvin_offset + 25
        Temp_C = B / np.log(Rth / (R25 * np.exp(-B / T25))) - Kelvin_offset
        return Temp_C

    def pcb_built_in_self_test(self):
        resp = self.gme2.request(GME2.CMD_TEST, GME2.CMD_GET_BIST_FLAGS, 1)
        return resp[2]

    def lcd_test_pattern(self, pattern):
        resp = self.gme2.request(GME2.CMD_TEST, GME2.CMD_LCD_TEST_PATTERN, pattern)
        return resp[3]

    def get_key_op(self):
        resp = self.gme2.request(GME2.CMD_TEST, GME2.CMD_GET_KEY)
        return resp[2], resp[3]

    def wait_key_press(self, timeout_sec):
        timeout = time.clock() + timeout_sec
        while time.clock() < timeout:
            key, op = self.get_key_op()
            if op == self.gme2.KEY_PRESS:
                return key
        return None

    def set_channel(self, channel):
        self.gme2.request(GME2.CMD_SET_CHANNEL, (channel >> 8) & 0xFF, channel & 0xFF)

    def set_frequency(self, frequency_Hz):
        new_freq = int(round((frequency_Hz / 1e6) / GME2.PLL_STEP))
        self.gme2.request(GME2.CMD_TEST, GME2.CMD_SET_FREQUENCY, (new_freq >> 8) & 0xFF, new_freq & 0xFF)
        self.frequency = new_freq * GME2.PLL_STEP

    def set_bandwidth(self, wide_bw):
        if wide_bw: # wide
            bw_bitmap = GME2.TEST_OPTIONS_WB_DEVIATION | GME2.TEST_OPTIONS_WB_FILTER
        else: # narrow
            bw_bitmap = 0
        self.gme2.request(GME2.CMD_TEST, GME2.CMD_TEST_OPTIONS, bw_bitmap)

    def set_rx_cal(self, dac, value):
        # Set RX calibration EEPROM and the DAC. Precondition: Frequency has been set via set_frequency().
        self.write_eeprom_byte(GME2.rx_cal_address[int(self.frequency)][dac], value)

    def get_rx_cal(self, dac):
        resp = self.read_eeprom_byte(GME2.rx_cal_address[int(self.frequency)][dac])
        return resp

    def set_tx_mod(self, value):
        # Set TX modulation EEPROM and the digital pot. Precondition: Frequency has been set via set_frequency().
        self.write_eeprom_byte(GME2.tx_mod_address[int(self.frequency)], value)

    def get_tx_mod(self):
        resp = self.read_eeprom_byte(GME2.tx_mod_address[int(self.frequency)])
        return resp

    def set_tx_bal(self, value):
        # Set TX balance EEPROM and the digital pot.
        value = ord(struct.pack("b", value)[0])
        self.write_eeprom_byte(GME2.EE_BAL, value)

    def get_tx_bal(self):
        resp = self.read_eeprom_byte(GME2.EE_BAL)
        resp = struct.unpack("b", chr(resp))[0]
        return resp

    def set_tx_aux_power_dac(self, value):
        # Set TX auxiliary PA power calibration EEPROM and the digital pot.
        # Precondition: Frequency has been set via set_frequency().
        self.write_eeprom_byte(GME2.tx_aux_power_cal_address[int(self.frequency)], value)

    def get_tx_aux_power_dac(self):
        resp = self.read_eeprom_byte(GME2.tx_aux_power_cal_address[int(self.frequency)])
        return resp

    def set_tx_power_dac(self, value):
        # Set TX auxiliary PA power calibration EEPROM and the digital pot.
        # Precondition: Frequency has been set via set_frequency().
        self.write_eeprom_byte(GME2.tx_power_cal_address[int(self.frequency)], value)

    def get_tx_power_dac(self):
        resp = self.read_eeprom_byte(GME2.tx_power_cal_address[int(self.frequency)])
        return resp

    def set_tx_power_scale_1w(self, value):
        # Set TX auxiliary PA power calibration EEPROM and the digital pot.
        # Precondition: Frequency has been set via set_frequency().
        self.write_eeprom_byte(GME2.EE_PWR_SCALE_1W, value)

    def get_tx_power_scale_1w(self):
        resp = self.read_eeprom_byte(GME2.EE_PWR_SCALE_1W)
        return resp

    def set_dac(self, dac, value):
        # Bypass the firmware logic and set the DAC directly.
        if dac not in GME2.DAC_channels:
            raise CommsError('Coding error - incorrect DAC channel.')
        self.gme2.request(GME2.CMD_TEST, GME2.CMD_WRITE_DAC, dac, value)

    def set_gpio(self, gpio, value):
        if gpio not in GME2.GPIO_ids:
            raise CommsError('Coding error - incorrect GPIO.')
        if value:
            value_byte = 1
        else:
            value_byte = 0
        self.gme2.request(GME2.CMD_TEST, GME2.CMD_WRITE_GPIO, gpio, value_byte)

    def set_pll_mux_out(self, value):
        self.gme2.request(GME2.CMD_TEST, GME2.CMD_WRITE_PLLMUX, value)

    def get_pll_unlock(self):
        resp = self.gme2.request(GME2.CMD_TEST, GME2.CMD_READ_PLL_UNLOCK)
        return resp[2]

    def get_freq_error(self):
        # Read frequency error. This is just the averaged voltage from RX audio ADC.
        resp = self.gme2.request(GME2.CMD_TEST, GME2.READ_FREQ_ERROR)
        # Convert the frequency error to signed integer
        return struct.unpack('>h', resp[2:4])[0]

    def set_pll_fref(self, value):
        # Adjust PLL reference (TCXO)
        self.write_eeprom_byte(GME2.EE_PLL_FREF, value)

    def get_pll_fref(self):
        resp = self.read_eeprom_byte(GME2.EE_PLL_FREF)
        return resp

    def set_pll_freq(self, frequency):
        # Write PLL frequency directly (in Hz)
        self.gme2.request(GME2.CMD_TEST, GME2.CMD_WRITE_PLL_FREQ,
            (frequency >> 24) & 0xFF,
            (frequency >> 16) & 0xFF,
            (frequency >> 8) & 0xFF,
            (frequency >> 0) & 0xFF)

    def get_pll_freq(self):
        resp = self.gme2.request(GME2.CMD_TEST, GME2.CMD_READ_PLL_FREQ)
        return (resp[2] << 24) | (resp[3] << 16) | (resp[4] << 8) | resp[5]

    def tx_on(self, pa_selection, modulation_selection, modulation_tone):
        # Set the power: 0 - 0.1W, 1 - 1W, 2 - 5W
        if pa_selection == 0: # 0.1W
            self.gme2.request(GME2.CMD_TEST, GME2.CMD_SELECT_PA, 0) # select aux PA
            self.gme2.request(GME2.CMD_TEST, GME2.CMD_SET_TX_PWR_LEVEL, 1) # do not scale
        elif pa_selection == 1: # 1W
            self.gme2.request(GME2.CMD_TEST, GME2.CMD_SELECT_PA, 1) # select main PA
            self.gme2.request(GME2.CMD_TEST, GME2.CMD_SET_TX_PWR_LEVEL, 0) # scale down to 1W
        elif pa_selection == 2: # 5W
            self.gme2.request(GME2.CMD_TEST, GME2.CMD_SELECT_PA, 1) # select main PA
            self.gme2.request(GME2.CMD_TEST, GME2.CMD_SET_TX_PWR_LEVEL, 1) # full scale at 5W
        # Press the virtual PTT
        self.gme2.request(GME2.CMD_SET_PTT, 1)
        # Setup the modulation source: 0 - no modulation, 1 - mike, 2 - tone generator
        if modulation_selection == 0:
            self.gme2.request(GME2.CMD_TEST, GME2.CMD_SET_TX_TONE, 0, 0, 0)
        elif modulation_selection == 1:
            pass # default - do nothing, use the mike
        else:
            self.gme2.request(GME2.CMD_TEST, GME2.CMD_SET_TX_TONE, 1, (modulation_tone >> 8) & 0xFF, modulation_tone & 0xFF)

    def set_mod_tone(self, mod_tone):
        self.gme2.request(GME2.CMD_TEST, GME2.CMD_SET_TX_TONE, 1, (mod_tone >> 8) & 0xFF, mod_tone & 0xFF)

    def tx_off(self):
        self.gme2.request(GME2.CMD_SET_PTT, 0)

    def dsp_test_capture(self, source):
        if source not in GME2.DTC_ids:
            raise CommsError('Coding error - DSP source ID is incorrect.')
        self.gme2.request(GME2.CMD_TEST, GME2.DSP_TEST_CAPTURE_START, source & 0xFF)
        buf_size = 600
        block_size = 100
        expected_resp_length = 2 * block_size + 3
        samples = np.zeros(buf_size, np.int16)
        for i in range(buf_size/block_size):
            timeout = time.clock() + 0.5 # expecting ~ 0.1 seconds at 8kHz
            while True:
                time.sleep(0.01)
                if time.clock() > timeout:
                    raise CommsTimeout('DSP test buffering timeout')
                try:
                    resp = self.gme2.request(GME2.CMD_TEST, GME2.DSP_TEST_CAPTURE_READ, i)
                except CommsTimeout:
                    continue
                if len(resp) == 6 and resp[2] == 0xFF:  # capture running, no data yet
                    continue
                elif len(resp) == expected_resp_length and resp[2] == i: # normal data
                    break
                else: # corrupted data
                    raise CommsError('Comms error - DSP test buffering issue (%d)' % len(resp))
            samples[i * block_size : (i + 1) * block_size] = struct.unpack('<%dh' % block_size, resp[3:])
        return samples

    def get_bt_mac(self):
        self.port.write('\nat+goi\n')
        time.sleep(0.1)
        resp = self.port.read(1024)
        i = resp.find('+GOI:')
        if i < 0:
            raise CommsError('Invalid Bluetooth MAC address response')
        s = resp[i+6:i+23].replace(':', '')
        return s

if __name__ == '__main__':
    # Unit test
    # Use default COM port or specify via command line argument
    if len(sys.argv) > 1:
        com_port = sys.argv[1]
    else:
        com_port = 'COM9'
    log = mtl.get_logger('cx50_control', True, mtl.logging.INFO)
    try:
        radio = CX50(com_port)
        log.info('Battery ADC: {}'.format(radio.adc_batt_volts()))
        log.info('RSSI ADC: {}'.format(radio.rssi_adc()))
        log.info('Noise ADC: {}'.format(radio.noise_adc()))
        log.info('Mike ADC: {}'.format(radio.mike_adc()))
        log.info('RX ADC: {}'.format(radio.rx_adc()))
        log.info('Top keys ADC: {}'.format(radio.adc_keys_top()))
        log.info('Bottom keys ADC: {}'.format(radio.adc_keys_bottom()))
        log.info('QB keys ADC: {}'.format(radio.adc_keys_qb()))
        radio.set_channel(0)
        for freq in 472, 472.1, 472.2:
            radio.set_frequency(freq)
            time.sleep(0.1)
            log.info('Frequency: {}, RSSI: {}'.format(freq, radio.rssi_adc()))
        log.info('Test completed.')
    except CommsTimeout:
        log.info('No response from the radio.')
    except:
        tb = traceback.format_exc()
        log.error(tb)
