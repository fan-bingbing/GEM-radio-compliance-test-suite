#!/usr/bin/python

'''
Top level manufacturing test script for CX50 Bed of Nails (PCB test)

Documentation: http://wiki.gme.net.au/display/PDLC/CX50+Manufacturing+Software

Copyright (C) 2018 Standard Communications Pty Ltd (GME). All rights reserved.
'''

import numpy as np

# We need this for delays
import time

# Serial port
import serial

# For command-line arguments
import sys

# Communications with the radio over the UART (programming cable) using GME2
import cx50_control

# Firmware flashing
import cx50_rx113_flash

# Communications with the test equipment
import cx50_manufacturing_test_equipment

# Test GUI
import cx50_manufacturing_test_gui

# The database
import cx50_manufacturing_test_database

# SoundCard audio in/out
import cx50_manufacturing_test_audio

# FFT windows
import fft_window

import matplotlib.pyplot as plt

import platform

if platform.system() == 'Linux':
    import cx50_manufacturing_test_printer as printer
    use_printer = True
else:
    use_printer = False

# Select com port
if len(sys.argv) > 1:
    com_port = sys.argv[1]
elif platform.system() == 'Linux':
    com_port = '/dev/ttyO4'
else:
    com_port = 'COM7'

if 'ttyO' in com_port:  # Is it BBB native UART
    import Adafruit_BBIO.UART as UART
    import Adafruit_BBIO.GPIO as GPIO
    UART.setup('UART' + com_port[-1])
    gpio_ptt = 'P9_15'  # GPIO_48
    GPIO.setup(gpio_ptt, GPIO.OUT)

# print text styles
text_reset = '\033[0m'
text_fg_red = '\033[91m'
text_fg_green = '\033[92m'
text_fg_yellow = '\033[93m'


# Container for constants / magic numbers
class Constants(object):
    tx_on_frequency = 476e6  # Hz  It seems easier to turn tx on at this frequency due to squelch logic
    single_align_frequency = 477e6  # Hz  This is used for frequency alignment and bal alignment
    multi_align_frequencies = 450e6, 477e6, 504e6, 520e6  # Hz
    rx_bpf_align_frequencies = 403e6, 450e6, 477e6, 504e6, 520e6  # Hz
    harmonics_test_frequency = 450e6  # Hz

    power_supply_voltage = 8.0  # V
    initial_current_min = 0.02  # A
    initial_current_max = 0.15  # A
    pre_fw_current_min = 0.02  # A
    pre_fw_current_max = 0.10  # A
    post_fw_current_min = 0.04  # A
    post_fw_current_max = 0.15  # A
    tx_off_current_max = 0.30  # A

    # tx current limits: 0.1W, 1W, 5W
    tx_on_current_min = [0.1, 0.2, 1.0]  # A
    tx_on_current_max = [0.3, 1.5, 2.0]  # A
    tx_on_stable_time = [1.0, 2.0, 3.0]  # Seconds
    tx_on_attenuation = [10,  20,  30]  # dB

    frequency_error_min = 25  # Hz
    frequency_tolerance = 50  # Hz
    tx_power_tolerance = 0.1  # dB
    aux_tx_power_target = 20  # dBm
    pri_tx_power_targets = [36, 37, 37, 36]  # dBm
    pri_tx_power_target_1w = 30  # dBm

    fm_deviation_target = 2500  # Hz
    fm_deviation_tolerance = 50  # Hz
    tx_mod_align_tone = 1000  # Hz
    tx_bal_align_tone = 100  # Hz
    rx_tune_tone = 1000  # Hz
    rx_tune_fm_dev = 1500  # Hz

    # Alignment feedback loop gains
    freq_align_loop_gain = 0.015  # frequency alignment loop gain (0.01 too slow, 0.02 unstable)
    aux_pa_align_loop_gain = 20  # secondary pa alignment loop gain (10 too slow, 30 unstable)
    pri_pa_align_loop_gain = 12  # primary pa alignment loop gain (5 too slow, 20 unstable)
    pri_pa_scale_loop_gain = 7  # primary pa scale loop gain (5 too slow, 20 unstable)
    tx_mod_align_loop_gain = 0.02  # tx mod alignment loop gain (0.01 too slow, 0.04 unstable)
    tx_bal_align_loop_gain = 0.03  # tx bal alignment loop gain (0.02 too slow, 0.05 unstable)

    max_align_loops = 10  # maximum alignment feedback loops

    # RX tuning start values:  T1, T2, T3, these should be typical values minus 25 because tuning span is 50.
    rx_tune_bases = np.array([[0, 5, 25],  # 403MHz
                              [20, 40, 70],  # 450MHz
                              [50, 70, 100],  # 477MHz
                              [80, 90, 130],  # 504MHz
                              [100, 100, 150]])  # 520MHz

    rx_tune_power = -90  # dBm
    rx_tuned_rssi_min = 100  # ADC
    rx_tuned_rssi_max = 150  # ADC

    # Sanity check limits and defaults
    sanity_fref_min = 1
    sanity_fref_max = 254
    sanity_fref_def = 130
    sanity_aux_pa_min = 100
    sanity_aux_pa_max = 160
    sanity_aux_pa_def = 130
    sanity_pri_pa_min = 100
    sanity_pri_pa_max = 160
    sanity_pri_pa_def = 130
    sanity_pri_pa_1w_min = 100
    sanity_pri_pa_1w_max = 160
    sanity_pri_pa_1w_def = 130
    sanity_tx_mod_min = 5
    sanity_tx_mod_max = 125
    sanity_tx_mod_def = 65
    sanity_tx_bal_min = 5
    sanity_tx_bal_max = 125
    sanity_tx_bal_def = 45

    # VCO range test settings
    vco_test_min_offset_on = 390e6  # Hz
    vco_test_max_offset_on = 500e6  # Hz
    vco_test_min_offset_off = 430e6  # Hz
    vco_test_max_offset_off = 550e6  # Hz
    vco_test_step = 1e6  # Hz
    vco_lock_min_offset_on = 410e6  # Hz
    vco_lock_max_offset_on = 450e6  # Hz
    vco_lock_min_offset_off = 450e6  # Hz
    vco_lock_max_offset_off = 520e6  # Hz

    # SINAD and sensitivity test
    sinad_source = 'SoundCard'
    # sinad_source = cx50_control.GME2.DTC_RX_ADC
    if sinad_source == 'SoundCard':
        sinad_fs = 44100  # sample rate in Hz
        sinad_n_samples = 1024  # number of samples
    else:
        sinad_fs = 32000  # sample rate in Hz
        sinad_n_samples = 600  # number of samples
    sinad_tone = 1000  # Hz
    sinad_fm_dev = 1500  # Hz
    sinad_min_freq = 300  # Hz
    sinad_max_freq = 3400  # Hz
    sinad_test_power = -120  # dBm

    # S1 & S10 calibration settings
    s1_power_level = -121  # dBm
    s10_power_level = -74  # dBm
    sanity_s1_min = 40
    sanity_s1_max = 100
    sanity_s10_min = 140
    sanity_s10_max = 200

    # noise level calibration settings
    noise_cal_frequency = 478e6  # Hz
    noise_cal_target = 112
    noise_scale_min = 200
    noise_scale_max = 400

    # audio volume adjustment settings
    audio_volume_min = 0.9
    audio_volume_max = 1.1

    # mic jack ptt test settings
    mic_ptt_test_tones = [0, 300, 1000, 3000]  # Hz
    mic_ptt_fm_dev_min = [50, 2000, 2000, 2000]  # Hz
    mic_ptt_fm_dev_max = [300, 2600, 2600, 2600]  # Hz
    mic_ptt_tx_power_min = 36.5  # dBm
    mic_ptt_tx_power_max = 37.5  # dBm

# A trivial data container for test results
class TestResults(object):
    def __init__(self):
        self.passed = True
        self.unique_id = 'unknown'
        self.serial_no = 'failed'
        self.failure_report_message = None
        self.log_list = []

    def add_line(self, line, to_print=True):
        if to_print:
            print line
        self.log_list.append(line)

    def save_log(self):
        with open('log/%s_%s.log' % (self.serial_no, self.unique_id), 'a+') as f:
            for item in self.log_list:
                f.write(item + '\n')

class Tester(object):
    def __init__(self, com_port):
        # Initialize the user interface
        self.gui = cx50_manufacturing_test_gui.init()
        self.gui.text('Starting CP50 MTS...')
        # Check the database connection (I assume this won't ever be the BeagleBone)
        self.gui.text('Initializing database...')
        error_msg = cx50_manufacturing_test_database.init()
        if error_msg:
            # This call doesn't return. User is required to fix the database connection and reboot the BeagleBone.
            self.halt_on_error(error_msg)
        # Check the test setup (self test)
        self.gui.text('Initializing test equipment...')
        error_msg = cx50_manufacturing_test_equipment.init()
        # Provide short alias names for the test equipment (we use them everywhere)
        self.power_supply = cx50_manufacturing_test_equipment.power_supply
        self.rf_attenuator = cx50_manufacturing_test_equipment.rf_attenuator
        self.rf_transceiver = cx50_manufacturing_test_equipment.rf_transceiver
        if error_msg:
            # This call doesn't return. User is required to fix the test equipment and reboot the BeagleBone.
            self.halt_on_error(error_msg)
        try:
            self.com_port = serial.Serial(port=com_port, baudrate=9600, timeout=0)
        except:
            self.halt_on_error('Failed to open serial port.')
        # Turn hardware PTT off by default
        self.hardware_ptt(False)
        # Initialize sound card
        try:
            self.audio = cx50_manufacturing_test_audio.SoundCard()
        except:
            self.halt_on_error('Failed to find sound card.')
        # Initialize label printer
        if use_printer:
            error_msg = printer.init()
            if error_msg:
                self.halt_on_error(error_msg)
        # Load firmware files to memory
        self.rx_flash = cx50_rx113_flash.RxFlash(self.com_port)
        self.is_mot_ok = self.rx_flash.read_mot_file('CP50.mot')
        self.is_geb_ok = self.rx_flash.read_geb_file('CP50.geb')
        if not self.is_mot_ok:
            self.halt_on_error()
        self.radio = cx50_control.CX50(self.com_port)
        self.sinad_fft_w = fft_window.Window(Constants.sinad_n_samples, 'HFT95')
        self.build_sinad_masks()
        self.skip_fw_load = False
        # Load last serial number of this test station
        try:
            self.sn_file = open('log/_sn_id.log', 'a+')
            all_lines = self.sn_file.readlines()
            last_line = all_lines[-1]
            self.test_station = last_line[4]
            self.last_sn_date = last_line[0:4]
            self.last_sn_count = int(last_line[5:9])
            print 'Last SN =', last_line[0:9]
        except:
            self.halt_on_error('Failed to find last SN.')

    def hardware_ptt(self, on=True):
        if 'ttyO' in com_port:
            # Use GPIO to control hardware PTT if using BBB native UART
            if on:
                GPIO.output(gpio_ptt, GPIO.HIGH)
            else:
                GPIO.output(gpio_ptt, GPIO.LOW)
        else:
            # Use RTS to control hardware PTT if using USB serial cable
            if on:
                self.com_port.setRTS(0)  # logic level flipped, RTS = HIGH
            else:
                self.com_port.setRTS(1)  # logic level flipped, RTS = LOW

    def run(self):
        # We call the tests from a sequential list and exit as soon as a test fails.
        test_sequence = (
            self.radio_test_start,
            self.prepare_fw_load,
            self.load_firmware,
            self.reboot_and_check_current_consumption,
            self.check_radio_serial_comms,
            self.get_radio_unique_id,
            self.radio_built_in_self_test,
            self.check_pcb_temperature,
            self.check_battery_adc,
            # self.check_rssi_adc,
            # self.check_noise_voltage_adc,
            # self.check_mike_adc,
            # self.check_rx_adc,
            # self.radio_lcd_test,
            # self.radio_keys_test,
            # self.frequency_alignment,
            #self.vco_range_test,
            #self.secondary_pa_bias_alignment,
            #self.tx_mod_bal_alignment,
            # self.tx_p25_mod_fidelity,
            #self.primary_pa_bias_alignment,
            # self.mic_ptt_test,
            # self.primary_pa_harmonics_test,
            # self.secondary_pa_harmonics_test,
            # self.secondary_trx_chip_test,
            #self.rx_filter_alignment,
            #self.s1_s10_calibration,
            #self.noise_level_calibration,
            # self.check_audio_volume,
            # self.rx_frequency_alignment,
            self.rx_sensitivity_test,
            # self.bluetooth_rf_test,
            # self.if_filter_sweep,  # RSSI and ADC average ("S curve") in +/-10kHz, TBD which steps
            self.print_label,
        )
        # TBD - ability to run individual tests vs the whole sequence..
        # Repeat tests for each radio forever (until the BeagleBone is shut off)
        while True:
            # New set of test results for each radio
            self.test_results = TestResults()
            # Radio tx is off by default
            self.is_radio_tx_on = False
            # Perform tests in sequence
            for test in test_sequence:
                try:
                    test()
                except:
                    self.test_results.failure_report_message = 'Exception: ' + str(sys.exc_info()[0])
                    self.test_results.passed = False
                if not self.test_results.passed:
                    break
            self.finalise_test()

    def radio_test_start(self):
        # Initial test equipment state
        self.power_supply.voltage = Constants.power_supply_voltage
        self.power_supply.current_limit = Constants.initial_current_max
        self.power_supply.on = False

        self.test_results.add_line('---------- Test Equipment ----------', False)
        self.test_results.add_line('power_supply_idn = ' + self.power_supply.idn_str, False)
        self.test_results.add_line('rf_attenuator_idn = ' + self.rf_attenuator.idn_str, False)
        self.test_results.add_line('rf_transceiver_idn = ' + self.rf_transceiver.idn_str, False)

        self.test_results.add_line('---------- Starting Test ----------')
        self.gui.message('Please plug in the radio, turn it on, then press Enter.')
        self.test_start_time = time.time()
        self.test_results.add_line('test_start_time = ' + time.strftime('%Y-%m-%d %H:%M:%S'))

        # Turn on the power and check the initial current
        self.power_supply.on = True
        time.sleep(1)
        current = self.power_supply.current
        self.test_results.add_line('initial_current = %.3f A' % current)
        self.assert_within_range('initial_current', current, Constants.initial_current_min, Constants.initial_current_max)
        self.power_supply.on = False

    def prepare_fw_load(self):
        self.skip_fw_load = not self.gui.question('Load firmware?')
        if self.skip_fw_load:
            return
        self.power_supply.on = False
        self.gui.message('Please hold radio buttons F1 and UP, then press Enter.')
        self.power_supply.on = True
        time.sleep(1)
        self.gui.message('Please release radio buttons, then press Enter.')
        # Expecting certain power consumption at this stage (narrow tolerance on this one.. the CPU is stopped, etc.)
        current = self.power_supply.current
        self.test_results.add_line('pre_fw_current = %.3f A' % current)
        self.assert_within_range('pre_fw_current', current, Constants.pre_fw_current_min, Constants.pre_fw_current_max)

    def power_cycle(self):
        self.power_supply.on = False
        time.sleep(0.2)
        self.power_supply.on = True

    def load_firmware(self):
        if self.skip_fw_load:
            return
        # Try sending the firmware image
        self.test_results.add_line('---------- Flashing Firmware ----------')
        start_time = time.time()
        # Program mot file
        result = self.rx_flash.program_mot_file()
        self.assert_true('flash_mot_ok', result)
        if not result:
            return  # failed to flash mot file
        self.test_results.add_line('flash_mot_time = %d seconds' % (time.time() - start_time))
        if not self.is_geb_ok:
            return  # geb file not found, skip
        start_time = time.time()
        # Reboot radio
        self.power_cycle()
        time.sleep(1)
        # Program geb file
        result = self.rx_flash.program_geb_file()
        self.assert_true('flash_geb_ok', result)
        if result:
            self.test_results.add_line('flash_geb_time = %d seconds' % (time.time() - start_time))

    def reboot_and_check_current_consumption(self):
        self.test_results.add_line('---------- Reboot and Check ----------')
        # The firmware has been loaded. Reboot and check the current consumption.
        self.power_supply.voltage = Constants.power_supply_voltage
        self.power_supply.current_limit = Constants.post_fw_current_max
        self.power_cycle()
        time.sleep(3)
        current = self.power_supply.current
        self.test_results.add_line('post_fw_current = %.3f A' % current)
        self.assert_within_range('post_fw_current', current, Constants.post_fw_current_min, Constants.post_fw_current_max)

    def check_radio_serial_comms(self):
        try:
            self.radio.check_connection()
            self.radio.write_eeprom_byte(self.radio.gme2.EE_SQL_LEVEL, 0)  # open squelch for subsequent tests
            self.test_results.add_line('fw_version = %s' % self.radio.firmware_version_string)
            comms_ok = True
        except:
            comms_ok = False
        self.assert_true('radio_serial_comms_ok', comms_ok)

    def get_radio_unique_id(self):
        # Use Bluetooth MAC address as PCB ID
        self.test_results.unique_id = self.radio.get_bt_mac()
        self.test_results.add_line('unique_id = ' + self.test_results.unique_id)

    def radio_built_in_self_test(self):
        # Firmware checks that all the on-board peripherals respond)
        error = self.radio.pcb_built_in_self_test()
        self.test_results.add_line('radio_bist_error = 0x%X' % error)
        self.assert_within_range('radio_bist_error', error, 0, 0)

    def check_pcb_temperature(self):
        # The tests and alignment are only valid if the temperature is within certain range.
        # (Using the temperature sensor ADC (NTC thermistor) on the PCB)
        pcb_temp = self.radio.adc_temperature()
        self.test_results.add_line('pcb_temperature = %.1f C' % pcb_temp)
        self.assert_within_range('pcb_temperature', pcb_temp, 0, 50)

    def check_battery_adc(self):
        batt_volt = self.radio.adc_batt_volts()
        self.test_results.add_line('battery_voltage = %.2f V' % batt_volt)
        self.assert_within_range('battery_voltage', batt_volt,
                                 Constants.power_supply_voltage - 0.3,
                                 Constants.power_supply_voltage + 0.3)

    def check_rssi_adc(self):
        self.assert_within_range('rssi_adc', self.radio.rssi_adc(), 10, 240)

    def check_noise_voltage_adc(self):
        self.assert_within_range('noise_adc', self.radio.noise_adc(), 10, 240)

    def check_mike_adc(self):
        self.assert_within_range('mike_adc', self.radio.mike_adc(), 10, 240)

    def check_rx_adc(self):
        self.assert_within_range('rx_adc', self.radio.rx_adc(), 10, 240)

    def radio_lcd_test(self):
        self.test_results.add_line('---------- LCD Test ----------')
        self.radio.lcd_test_pattern(0)
        answer = self.gui.question('All segments off?')
        self.assert_true('lcd_all_segments_off', answer)
        if not self.test_results.passed:
            return
        self.radio.lcd_test_pattern(1)
        answer = self.gui.question('All segments on?')
        self.assert_true('lcd_all_segments_on', answer)
        if not self.test_results.passed:
            return
        self.radio.lcd_test_pattern(2)
        answer = self.gui.question('Character segments ok?')
        self.assert_true('lcd_character_segments_ok', answer)
        if not self.test_results.passed:
            return
        self.radio.lcd_test_pattern(3)
        answer = self.gui.question('Icon segments ok?')
        self.assert_true('lcd_icon_segments_ok', answer)
        if not self.test_results.passed:
            return
        self.radio.lcd_test_pattern(4)
        answer = self.gui.question('Left segments ok?')
        self.assert_true('lcd_left_segments_ok', answer)
        if not self.test_results.passed:
            return
        self.radio.lcd_test_pattern(5)
        answer = self.gui.question('Right segments ok?')
        self.assert_true('lcd_right_segments_ok', answer)
        if self.test_results.passed:
            self.test_results.add_line('LCD test passed')

    def radio_keys_test(self):
        self.test_results.add_line('---------- Keys Test ----------')
        key_names = ['F1', 'F2', 'F3', 'F4', 'UP', 'MENU', 'DOWN', 'QB1', 'QB2', 'QB3', 'EMERGENCY']
        key_codes = [self.radio.gme2.F1_KEY_H, self.radio.gme2.F2_KEY_H, self.radio.gme2.F3_KEY_H, self.radio.gme2.F4_KEY_H,
                     self.radio.gme2.UP_KEY_H, self.radio.gme2.A_KEY_H, self.radio.gme2.DN_KEY_H,
                     self.radio.gme2.QB1_KEY_H, self.radio.gme2.QB2_KEY_H, self.radio.gme2.QB3_KEY_H,
                     self.radio.gme2.EMERG_KEY_H]
        self.gui.text(text_fg_yellow + 'Please press ' + ', '.join(key_names) + ' keys in sequence' + text_reset)
        test_ok = True
        for i in range(len(key_codes)):
            key = self.radio.wait_key_press(10)  # timeout 10 seconds
            if key is None:
                self.gui.text(text_fg_red + 'Wait for key press timeout' + text_reset)
                test_ok = False
                break
            elif key == key_codes[i]:
                self.gui.text(key_names[i] + ' key ok')
            else:
                self.gui.text(text_fg_red + 'Unexpected key code: ' + str(key) + text_reset)
                test_ok = False
                break
        self.assert_true('all_keys_ok', test_ok)
        if not self.test_results.passed:
            return
        test_ok = False
        clockwise_ok = False
        self.gui.text(text_fg_yellow + 'Please turn channel knob clockwise' + text_reset)
        for i in range(10):  # maximum 10 knob clicks
            key = self.radio.wait_key_press(10)  # timeout 10 seconds
            if key is None:
                self.gui.text(text_fg_red + 'Wait for knob turning timeout' + text_reset)
                break
            elif key == self.radio.gme2.DN_KEY_H:
                if not clockwise_ok:
                    clockwise_ok = True
                    self.gui.text('Knob clockwise ok')
                    self.gui.text(text_fg_yellow + 'Please turn channel knob counter-clockwise' + text_reset)
            elif key == self.radio.gme2.UP_KEY_H:
                if clockwise_ok:
                    self.gui.text('Knob counter-clockwise ok')
                    test_ok = True
                    break
                else:
                    self.gui.text(text_fg_red + 'Please turn channel knob clockwise first' + text_reset)
            else:
                self.gui.text(text_fg_red + 'Unexpected key code: ' + str(key) + text_reset)
                break
        self.assert_true('channel_knob_ok', test_ok)
        if self.test_results.passed:
            self.test_results.add_line('Keys test passed')

    # pa_select: 0 = 0.1W, 1 = 1W, 2 = 5W
    # mod_source: 0 = none, 1 = mic, 2 = tone
    # mod_tone: modulation tone frequency in Hz
    def radio_tx_on(self, pa_select, mod_source=0, mod_tone=0):
        self.radio.set_frequency(Constants.tx_on_frequency)
        self.rf_attenuator.attenuation = Constants.tx_on_attenuation[pa_select]
        self.power_supply.current_limit = Constants.tx_on_current_max[pa_select]
        self.radio.tx_on(pa_select, mod_source, mod_tone)
        self.is_radio_tx_on = True
        time.sleep(Constants.tx_on_stable_time[pa_select])  # wait radio tx warming up

    def radio_tx_off(self):
        if self.is_radio_tx_on:
            self.radio.tx_off()
            self.is_radio_tx_on = False
            time.sleep(0.5)  # wait radio tx turning off
        self.power_supply.current_limit = Constants.tx_off_current_max

    def frequency_alignment(self):
        self.test_results.add_line('---------- Frequency Alignment ----------')
        self.assert_true('n200_gps_locked', self.rf_transceiver.is_gps_locked())
        if not self.test_results.passed:
            return
        self.radio_tx_on(0)  # radio tx on at 0.1W, no modulation
        freq = Constants.single_align_frequency
        self.radio.set_frequency(freq)
        self.rf_transceiver.rx_freq = freq
        # get current fref
        fref = self.radio.get_pll_fref()
        # sanity check
        if fref < Constants.sanity_fref_min or fref > Constants.sanity_fref_max:
            fref = Constants.sanity_fref_def
        # start auto alignment
        freq_err = 0
        for i in range(Constants.max_align_loops):
            fref += int(round(freq_err * Constants.freq_align_loop_gain))
            # sanity check
            if fref < Constants.sanity_fref_min:
                fref = Constants.sanity_fref_min
            elif fref > Constants.sanity_fref_max:
                fref = Constants.sanity_fref_max
            self.radio.set_pll_fref(fref)
            time.sleep(0.1)
            freq_err = self.rf_transceiver.get_rx_freq_offset()
            if abs(freq_err) < Constants.frequency_error_min:
                break  # finish alignment if error is minimal
            elif abs(freq_err) < Constants.frequency_tolerance:
                # try adjusting fref by 1 to minimise the error
                fref_1 = (fref + 1) if (freq_err > 0) else (fref - 1)
                self.radio.set_pll_fref(fref_1)
                time.sleep(0.1)
                freq_err_1 = self.rf_transceiver.get_rx_freq_offset()
                if abs(freq_err_1) > abs(freq_err):
                    # restore previous fref if error becomes bigger
                    self.radio.set_pll_fref(fref)
                    time.sleep(0.1)
                    freq_err = self.rf_transceiver.get_rx_freq_offset()
                else:
                    # adopt new fref if error becomes smaller or same
                    freq_err = freq_err_1
                break  # finish alignment after error minimisation
        self.test_results.add_line('freq = %d MHz, fref = %d, freq_error = %d Hz' % (freq / 1e6, fref, freq_err))
        self.assert_within_range('freq_error', freq_err, -Constants.frequency_tolerance, Constants.frequency_tolerance)

    def get_vco_range(self, freq_min, freq_max, freq_step):
        lock_min = 0
        lock_max = 0
        lock_count = 0
        # set pll chip mux output as analog lock signal
        self.radio.set_pll_mux_out(self.radio.gme2.PLL_MUX_ALD)
        for freq in range(int(freq_min), int(freq_max), int(freq_step)):
            self.radio.set_pll_freq(freq)
            # wait till pll is locked or timeout
            for i in range(10):
                time.sleep(0.01)
                locked = (self.radio.get_pll_unlock() == 0)
                if locked:
                    break
            # check if pll is finally locked
            if locked:
                lock_count += 1
                # debounce: record lock_min if consistently locked for 10 steps
                if lock_count == 10:
                    lock_min = freq - freq_step * (lock_count - 1)
                lock_max = freq
            else:
                lock_count = 0
                # done sweeping if pll becomes unlocked after lock_min is found
                if lock_min != 0:
                    break
        return lock_min, lock_max

    def vco_range_test(self):
        self.test_results.add_line('---------- VCO Range Test ----------')
        self.radio_tx_off()
        # Sweep and get VCO range with VCO offset turned on
        self.radio.set_gpio(self.radio.gme2.GPIO_VCO_OFFSET, True)
        lock_min, lock_max = self.get_vco_range(Constants.vco_test_min_offset_on, Constants.vco_test_max_offset_on, Constants.vco_test_step)
        self.test_results.add_line('vco_range_offset_on: %d - %d MHz' % (lock_min / 1e6, lock_max / 1e6))
        self.assert_within_range('vco_range_min_offset_on', lock_min, Constants.vco_test_min_offset_on, Constants.vco_lock_min_offset_on)
        if not self.test_results.passed:
            return
        self.assert_within_range('vco_range_max_offset_on', lock_max, Constants.vco_lock_max_offset_on, Constants.vco_test_max_offset_on)
        if not self.test_results.passed:
            return
        # Sweep and get VCO range with VCO offset turned off
        self.radio.set_gpio(self.radio.gme2.GPIO_VCO_OFFSET, False)
        lock_min, lock_max = self.get_vco_range(Constants.vco_test_min_offset_off, Constants.vco_test_max_offset_off, Constants.vco_test_step)
        self.test_results.add_line('vco_range_offset_off: %d - %d MHz' % (lock_min / 1e6, lock_max / 1e6))
        self.assert_within_range('vco_range_min_offset_off', lock_min, Constants.vco_test_min_offset_off, Constants.vco_lock_min_offset_off)
        if not self.test_results.passed:
            return
        self.assert_within_range('vco_range_max_offset_off', lock_max, Constants.vco_lock_max_offset_off, Constants.vco_test_max_offset_off)

    def secondary_pa_bias_alignment(self):
        self.test_results.add_line('---------- Secondary PA Alignment ----------')
        self.radio_tx_on(0)  # radio tx on at 0.1W, no modulation
        for freq in Constants.multi_align_frequencies:
            self.radio.set_frequency(freq)
            self.rf_transceiver.rx_freq = freq
            # get current dac value
            dac = self.radio.get_tx_aux_power_dac()
            # sanity check
            if dac < Constants.sanity_aux_pa_min or dac > Constants.sanity_aux_pa_max:
                dac = Constants.sanity_aux_pa_def
            # start auto alignment
            power_err = 0
            for i in range(Constants.max_align_loops):
                dac += int(round(power_err * Constants.aux_pa_align_loop_gain))
                # sanity check
                if dac < Constants.sanity_aux_pa_min:
                    dac = Constants.sanity_aux_pa_min
                elif dac > Constants.sanity_aux_pa_max:
                    dac = Constants.sanity_aux_pa_max
                self.radio.set_tx_aux_power_dac(dac)
                time.sleep(0.1)
                power_dbm = self.rf_transceiver.get_rx_power()
                power_err = Constants.aux_tx_power_target - power_dbm
                if abs(power_err) < Constants.tx_power_tolerance:
                    break  # finish alignment if error is in tolerance
                if dac >= Constants.sanity_aux_pa_max and power_err > 0:
                    # something wrong, quit alignment and set to minimum value
                    self.radio.set_tx_aux_power_dac(Constants.sanity_aux_pa_min)
                    break
            self.test_results.add_line('freq = %d MHz, dac = %d, tx_power = %.2f dBm, current = %.3f A' % (freq / 1e6, dac, power_dbm, self.power_supply.current))
            self.assert_within_range('aux_tx_power_dbm', power_dbm,
                                     Constants.aux_tx_power_target - Constants.tx_power_tolerance,
                                     Constants.aux_tx_power_target + Constants.tx_power_tolerance)
            if not self.test_results.passed:
                break  # failed to adjust power in tolerance

    def tx_mod_bal_alignment(self):
        self.test_results.add_line('---------- TX Mod/Bal Alignment ----------')
        # Mod alignment
        self.radio_tx_on(0, 2, Constants.tx_mod_align_tone)  # radio tx on at 0.1W, modulation tone 1kHz
        for freq in Constants.multi_align_frequencies:
            self.radio.set_frequency(freq)
            self.rf_transceiver.rx_freq = freq
            # get current mod value
            mod = self.radio.get_tx_mod()
            # sanity check
            if mod < Constants.sanity_tx_mod_min or mod > Constants.sanity_tx_mod_max:
                mod = Constants.sanity_tx_mod_def
            # start auto alignment
            fm_dev_err = 0
            for i in range(Constants.max_align_loops):
                mod += int(round(fm_dev_err * Constants.tx_mod_align_loop_gain))
                # sanity check
                if mod < Constants.sanity_tx_mod_min:
                    mod = Constants.sanity_tx_mod_min
                elif mod > Constants.sanity_tx_mod_max:
                    mod = Constants.sanity_tx_mod_max
                self.radio.set_tx_mod(mod)
                time.sleep(0.2)
                fm_dev = self.rf_transceiver.get_rx_fm_dev()
                fm_dev_err = Constants.fm_deviation_target - fm_dev
                if abs(fm_dev_err) < Constants.fm_deviation_tolerance:
                    break  # finish alignment if error is in tolerance
            self.test_results.add_line('freq = %d MHz, mod = %d, fm_dev = %d Hz' % (freq / 1e6, mod, fm_dev))
            self.assert_within_range('mod_fm_deviation', fm_dev,
                                     Constants.fm_deviation_target - Constants.fm_deviation_tolerance,
                                     Constants.fm_deviation_target + Constants.fm_deviation_tolerance)
            if not self.test_results.passed:
                return  # failed to adjust fm deviation in tolerance
        # Bal alignment
        self.radio_tx_on(0, 2, Constants.tx_bal_align_tone)  # radio tx on at 0.1W, modulation tone 100Hz
        freq = Constants.single_align_frequency
        self.radio.set_frequency(freq)
        self.rf_transceiver.rx_freq = freq
        # get current bal value
        bal = self.radio.get_tx_bal()
        # sanity check
        if bal < Constants.sanity_tx_bal_min or bal > Constants.sanity_tx_bal_max:
            bal = Constants.sanity_tx_bal_def
        # start auto alignment
        fm_dev_err = 0
        for i in range(Constants.max_align_loops):
            bal += int(round(fm_dev_err * Constants.tx_bal_align_loop_gain))
            # sanity check
            if bal < Constants.sanity_tx_bal_min:
                bal = Constants.sanity_tx_bal_min
            elif bal > Constants.sanity_tx_bal_max:
                bal = Constants.sanity_tx_bal_max
            self.radio.set_tx_bal(bal)
            time.sleep(0.2)
            fm_dev = self.rf_transceiver.get_rx_fm_dev()
            fm_dev_err = Constants.fm_deviation_target - fm_dev
            if abs(fm_dev_err) < Constants.fm_deviation_tolerance:
                break  # finish alignment if error is in tolerance
        self.test_results.add_line('freq = %d MHz, bal = %d, fm_dev = %d Hz' % (freq / 1e6, bal, fm_dev))
        self.assert_within_range('bal_fm_deviation', fm_dev,
                                 Constants.fm_deviation_target - Constants.fm_deviation_tolerance,
                                 Constants.fm_deviation_target + Constants.fm_deviation_tolerance,)

    def tx_p25_mod_fidelity(self):
        pass

    def primary_pa_bias_alignment(self):
        self.test_results.add_line('---------- Primary PA Alignment ----------')
        # reset eeprom tx power values to default
        for address in self.radio.gme2.tx_power_cal_address.values():
            self.radio.write_eeprom_byte(address, Constants.sanity_pri_pa_def)
        self.radio_tx_on(2)  # radio tx on at 5W, no modulation
        # High power = 5W
        idx = 0
        for freq in Constants.multi_align_frequencies:
            self.radio.set_frequency(freq)
            self.rf_transceiver.rx_freq = freq
            # get current dac value
            dac = self.radio.get_tx_power_dac()
            # sanity check
            if dac < Constants.sanity_pri_pa_min or dac > Constants.sanity_pri_pa_max:
                dac = Constants.sanity_pri_pa_def
            # start auto alignment
            power_err = 0
            for i in range(Constants.max_align_loops):
                dac += int(round(power_err * Constants.pri_pa_align_loop_gain))
                # sanity check
                if dac < Constants.sanity_pri_pa_min:
                    dac = Constants.sanity_pri_pa_min
                elif dac > Constants.sanity_pri_pa_max:
                    dac = Constants.sanity_pri_pa_max
                self.radio.set_tx_power_dac(dac)
                time.sleep(0.1)
                power_dbm = self.rf_transceiver.get_rx_power()
                power_err = Constants.pri_tx_power_targets[idx] - power_dbm
                if abs(power_err) < Constants.tx_power_tolerance:
                    break  # finish alignment if error is in tolerance
                if dac >= Constants.sanity_pri_pa_max and power_err > 0:
                    # something wrong, quit alignment and set to minimum value
                    self.radio.set_tx_power_dac(Constants.sanity_pri_pa_min)
                    break
            self.test_results.add_line('freq = %d MHz, dac = %d, tx_power = %.2f dBm, current = %.3f A' % (freq / 1e6, dac, power_dbm, self.power_supply.current))
            self.assert_within_range('pri_tx_power_dbm', power_dbm,
                                     Constants.pri_tx_power_targets[idx] - Constants.tx_power_tolerance,
                                     Constants.pri_tx_power_targets[idx] + Constants.tx_power_tolerance)
            idx += 1
            if not self.test_results.passed:
                return  # failed to adjust power in tolerance
        # Low power = 1W
        self.radio_tx_on(1)  # radio tx on at 1W, no modulation
        freq = Constants.single_align_frequency
        self.rf_transceiver.rx_freq = freq
        self.radio.set_frequency(freq)
        # get current dac value
        dac = self.radio.get_tx_power_scale_1w()
        # sanity check
        if dac < Constants.sanity_pri_pa_1w_min or dac > Constants.sanity_pri_pa_1w_max:
            dac = Constants.sanity_pri_pa_1w_def
        # start auto alignment
        power_err = 0
        for i in range(Constants.max_align_loops):
            dac += int(round(power_err * Constants.pri_pa_scale_loop_gain))
            # sanity check
            if dac < Constants.sanity_pri_pa_1w_min:
                dac = Constants.sanity_pri_pa_1w_min
            elif dac > Constants.sanity_pri_pa_1w_max:
                dac = Constants.sanity_pri_pa_1w_max
            self.radio.set_tx_power_scale_1w(dac)
            time.sleep(0.1)
            power_dbm = self.rf_transceiver.get_rx_power()
            power_err = Constants.pri_tx_power_target_1w - power_dbm
            if abs(power_err) < Constants.tx_power_tolerance:
                break  # finish alignment if error is in tolerance
        self.test_results.add_line('freq = %d MHz, scale = %d, tx_power_1w = %.2f dBm, current = %.3f A' % (freq / 1e6, dac, power_dbm, self.power_supply.current))
        self.assert_within_range('pri_tx_power_1w_dbm', power_dbm,
                                 Constants.pri_tx_power_target_1w - Constants.tx_power_tolerance,
                                 Constants.pri_tx_power_target_1w + Constants.tx_power_tolerance)

    def mic_ptt_test(self):
        self.test_results.add_line('---------- Mic Jack PTT Test ----------')
        self.radio_tx_on(2, 1)  # radio tx on at 5W, use mic
        self.rf_transceiver.rx_freq = Constants.tx_on_frequency
        self.hardware_ptt(True)
        for i in range(len(Constants.mic_ptt_test_tones)):
            tone = Constants.mic_ptt_test_tones[i]
            if tone > 0:
                self.audio.play_tone(tone)
            else:
                self.audio.stop_play()
            time.sleep(1)
            fm_dev = self.rf_transceiver.get_rx_fm_dev()
            power = self.rf_transceiver.get_rx_power()
            self.test_results.add_line('tone = %d Hz, fm_dev = %d Hz, power = %.2f dBm' % (tone, fm_dev, power))
            self.assert_within_range('mic_ptt_fm_dev', fm_dev, Constants.mic_ptt_fm_dev_min[i], Constants.mic_ptt_fm_dev_max[i])
            if not self.test_results.passed:
                break
            self.assert_within_range('mic_ptt_tx_power', power, Constants.mic_ptt_tx_power_min, Constants.mic_ptt_tx_power_max)
            if not self.test_results.passed:
                break
        self.hardware_ptt(False)
        self.audio.stop_play()
        self.radio_tx_off()

    def primary_pa_harmonics_test(self):
        self.test_results.add_line('---------- Primary PA Harmonics ----------')
        self.radio_tx_on(2)  # radio tx on at 5W, no modulation
        tx_freq = int(Constants.harmonics_test_frequency)
        self.radio.set_frequency(tx_freq)
        for rx_freq in range(tx_freq, int(3e9), tx_freq):
            self.rf_transceiver.rx_freq = rx_freq
            power_dbm = self.rf_transceiver.get_rx_power()
            self.test_results.add_line('freq = %d MHz, power = %.2f dBm' % (rx_freq / 1e6, power_dbm))

    def secondary_pa_harmonics_test(self):
        self.test_results.add_line('---------- Secondary PA Harmonics ----------')
        self.radio_tx_on(0)  # radio tx on at 0.1W, no modulation
        tx_freq = int(Constants.harmonics_test_frequency)
        self.radio.set_frequency(tx_freq)
        for rx_freq in range(tx_freq, int(3e9), tx_freq):
            self.rf_transceiver.rx_freq = rx_freq
            power_dbm = self.rf_transceiver.get_rx_power()
            self.test_results.add_line('freq = %d MHz, power = %.2f dBm' % (rx_freq / 1e6, power_dbm))

    def secondary_trx_chip_test(self):
        pass

    def rx_filter_alignment(self):
        self.test_results.add_line('---------- RX Filter Alignment ----------')
        self.radio_tx_off()
        self.rf_transceiver.mod_freq = Constants.rx_tune_tone
        self.rf_transceiver.fm_dev = Constants.rx_tune_fm_dev
        f_idx = 0
        for freq in Constants.rx_bpf_align_frequencies:
            tune_base = Constants.rx_tune_bases[f_idx]
            f_idx += 1
            self.radio.set_frequency(freq)
            self.rf_transceiver.tx_freq = freq
            self.rf_transceiver.set_tx_power(Constants.rx_tune_power)
            self.rf_transceiver.tx_on = True  # set tx_on to update tx signal
            max_cal = np.empty(3, dtype='uint8')
            scores = np.empty(5, dtype='uint8')
            for t in range(3):  # 0 = T1, 1 = T2, 2 = T3
                for i in range(5):  # tuning steps
                    self.radio.set_rx_cal(t, i * 10 + tune_base[t])
                    time.sleep(0.1)
                    scores[i] = self.radio.rssi_adc()
                max_idx1 = np.argmax(scores)
                max_idx2 = len(scores) - 1 - np.argmax(np.flip(scores, 0))
                max_cal[t] = (max_idx1 + max_idx2) * 5 + tune_base[t]
                self.radio.set_rx_cal(t, max_cal[t])
                time.sleep(0.1)
            for t in range(3):  # 0 = T1, 1 = T2, 2 = T3
                for i in range(5):  # tuning steps
                    self.radio.set_rx_cal(t, max_cal[t] + i * 2 - 4)
                    time.sleep(0.1)
                    scores[i] = self.radio.rssi_adc()
                max_idx1 = np.argmax(scores)
                max_idx2 = len(scores) - 1 - np.argmax(np.flip(scores, 0))
                # assume one step more if max occurs at either boundary
                if max_idx1 == 0:
                    max_idx1 = -1
                if max_idx2 == len(scores) - 1:
                    max_idx2 = len(scores)
                max_cal[t] += (max_idx1 + max_idx2) - 4
                self.radio.set_rx_cal(t, max_cal[t])
                time.sleep(0.1)
            rssi = self.radio.rssi_adc()
            self.test_results.add_line('freq = %d MHz, [T1 T2 T3] = %s, rssi = %d' % (freq / 1e6, max_cal, rssi))
            self.assert_within_range('rx_tuned_rssi', rssi, Constants.rx_tuned_rssi_min, Constants.rx_tuned_rssi_max)
            if not self.test_results.passed:
                break  # failed to tune rx filter in tolerance
        self.rf_transceiver.tx_on = False

    def check_audio_volume(self):
        self.test_results.add_line('---------- Check Audio Volume ----------')
        self.radio_tx_off()
        self.radio.set_frequency(Constants.single_align_frequency)
        self.rf_transceiver.tx_freq = Constants.single_align_frequency
        self.rf_transceiver.set_tx_power(Constants.rx_tune_power)
        self.rf_transceiver.mod_freq = Constants.rx_tune_tone
        self.rf_transceiver.fm_dev = Constants.rx_tune_fm_dev
        self.rf_transceiver.tx_on = True
        for i in range(100):  # timeout approx 10 seconds
            time.sleep(0.1)
            volume = self.get_mic_level()
            if volume >= Constants.audio_volume_min:
                break
            if i == 0:
                self.gui.text(text_fg_yellow + 'Please turn volume knob to maximum' + text_reset)
        self.rf_transceiver.tx_on = False
        self.test_results.add_line('audio_volume = %.2f' % volume)
        self.assert_within_range('audio_volume', volume, Constants.audio_volume_min, Constants.audio_volume_max)

    def rx_frequency_alignment(self):
        print '---------- RX Frequency Alignment ----------'
        self.radio_tx_off()
        freq = 477.5e6
        self.radio.set_frequency(freq)
        self.rf_transceiver.tx_freq = freq
        self.rf_transceiver.set_tx_power(Constants.rx_tune_power)
        self.rf_transceiver.mod_freq = Constants.rx_tune_tone
        self.rf_transceiver.fm_dev = Constants.rx_tune_fm_dev
        self.rf_transceiver.tx_on = True
        tx_fref = self.radio.get_pll_fref()  # get current fref
        best_fref = tx_fref
        best_sinad = 0
        for i in range(10, 31):
            rx_fref = tx_fref + i
            if rx_fref < Constants.sanity_fref_min:
                continue
            if rx_fref > Constants.sanity_fref_max:
                break
            self.radio.set_pll_fref(rx_fref)
            time.sleep(0.1)
            sinad = self.get_rx_sinad()
            if sinad > best_sinad:
                best_sinad = sinad
                best_fref = rx_fref
            print 'Fref = %d, SINAD = %d' % (rx_fref, sinad)
        self.rf_transceiver.tx_on = False
        self.radio.set_pll_fref(best_fref)
        print 'Best RX Fref = %d' % best_fref

    def if_filter_sweep(self):
        print '---------- IF Filter Sweep ----------'
        self.radio_tx_off()
        self.rf_transceiver.mod_freq = Constants.rx_tune_tone
        self.rf_transceiver.fm_dev = Constants.rx_tune_fm_dev
        legend_list = []
        x = np.arange(-15, 16)  # sweep frequency range in kHz
        rssi = np.empty(len(x), dtype='uint8')
        test_freqs = np.arange(499e6, 501.01e6, 0.1e6)
        count = 0
        for freq in test_freqs:
            legend_list.append('%.1f MHz' % (freq / 1e6))
            self.radio.set_frequency(freq)
            for i in range(len(x)):
                self.rf_transceiver.tx_freq = freq + x[i] * 1e3
                self.rf_transceiver.set_tx_power(Constants.rx_tune_power)
                self.rf_transceiver.tx_on = True  # set tx_on to update tx signal
                rssi[i] = self.radio.rssi_adc()
            print 'Freq = %.1f MHz ,' % (freq / 1e6), 'RSSI =', rssi
            if count < 10:
                style = '-'
            elif count == 10:
                style = '--'
            else:
                style = ':'
            plt.plot(x, rssi, style)
            count += 1
        self.rf_transceiver.tx_on = False
        plt.ylabel('RSSI')
        plt.xlabel('Frequency offset[kHz]')
        plt.legend(legend_list, loc='upper left')
        plt.grid()
        plt.show()

    def build_sinad_masks(self, weighted=True):
        tone_bin = int(Constants.sinad_n_samples * Constants.sinad_tone / Constants.sinad_fs)
        notch_width = 14
        tone_notch_min_bin = tone_bin - notch_width / 2
        tone_notch_max_bin = tone_bin + notch_width / 2
        sinad_min_freq_bin = int(Constants.sinad_n_samples * Constants.sinad_min_freq / Constants.sinad_fs)
        sinad_max_freq_bin = int(Constants.sinad_n_samples * Constants.sinad_max_freq / Constants.sinad_fs) + notch_width
        self.sinad_mask_db = np.zeros(Constants.sinad_n_samples / 2)
        self.sinad_mask_db[:sinad_min_freq_bin] = -120
        self.sinad_mask_db[sinad_max_freq_bin:] = -120
        if weighted:
            preemph_max_att = -6 * np.log2(float(sinad_max_freq_bin) / float(sinad_min_freq_bin))
            self.sinad_mask_db[sinad_min_freq_bin: sinad_max_freq_bin] = np.linspace(0, preemph_max_att, sinad_max_freq_bin - sinad_min_freq_bin)
        self.sinad_mask = 10 ** (self.sinad_mask_db / 10.0)
        self.sinad_mask_db[tone_notch_min_bin: tone_notch_max_bin] = -120
        self.sinad_mask_with_notch = 10 ** (self.sinad_mask_db / 10.0)

    def get_rx_sinad(self):
        if Constants.sinad_source == 'SoundCard':
            samples = self.audio.get_mic_samps(Constants.sinad_n_samples)
        else:
            samples = self.radio.dsp_test_capture(Constants.sinad_source)
        samples = 1.0 * samples - np.average(samples)  # Remove DC
        samples_fft_complex = np.fft.fft(samples * self.sinad_fft_w)[:Constants.sinad_n_samples / 2]
        samples_fft = samples_fft_complex.real ** 2 + samples_fft_complex.imag ** 2
        SND = np.sum(samples_fft * self.sinad_mask)
        ND = np.sum(samples_fft * self.sinad_mask_with_notch)
        SINAD = 10 * np.log10(SND / ND + 1e-10)

        plt.subplot(2, 1, 1)
        plt.plot(samples)
        plt.ylim(-0.4, 0.4)
        ax = plt.subplot(2, 1, 2)
        samples_fft_db = 10 * np.log10(samples_fft / len(samples) + 1e-10)
        plt.plot(samples_fft_db)
        plt.plot(self.sinad_mask_db)
        plt.text(0.4, 0.9, 'SINAD = %.2f dB' % SINAD, transform=ax.transAxes)
        plt.show()

        return SINAD

    def get_mic_level(self):
        if platform.system() == 'Linux':
            gain = 0.5
        else:
            gain = 3.4
        samples = self.audio.get_mic_samps(Constants.sinad_n_samples, gain)
        return np.ptp(samples)

    def rx_sensitivity_test(self):
        print '---------- RX Sensitivity Test ----------'
        f = open("sinad.csv", "w+")
        f.write('Frequency[Hz], Power[dBm], SINAD[dB]\n')
        self.radio_tx_off()
        sig_gen = cx50_manufacturing_test_equipment.signal_generator
        if sig_gen is not None:  # use sig gen
            print 'Use sig gen'
            self.rf_attenuator.attenuation = 80
            sig_gen.fm_dev = Constants.sinad_fm_dev
            sig_gen.mod_freq = Constants.sinad_tone
            sig_gen.level = -3.7
            sig_gen.fm_on = True
            sig_gen.mod_on = True
            sig_gen.rf_on = True
        else:  # use n200 tx
            print 'Use n200 tx'
            self.rf_transceiver.mod_freq = Constants.sinad_tone
            self.rf_transceiver.fm_dev = Constants.sinad_fm_dev
        test_freqs = np.arange(450e6, 520.01e6, 1e6)
        for freq in test_freqs:
            self.radio.set_frequency(freq)
            if sig_gen is not None:  # use sig gen
                sig_gen.frequency = freq
                power = sig_gen.level
                time.sleep(0.1)
            else:  # use n200 tx
                self.rf_transceiver.tx_freq = freq
                power = self.rf_transceiver.set_tx_power(Constants.sinad_test_power)
                self.rf_transceiver.tx_on = True  # set tx_on to update tx signal
            sinad = self.get_rx_sinad()
            print 'Freq = %d Hz, Power = %0.2f dBm, SINAD = %.2f dB' % (freq, power, sinad)
            f.write('%d, %0.2f, %.2f\n' % (freq, power, sinad))
        f.close()
        if sig_gen is not None:
            sig_gen.rf_on = False
        else:
            self.rf_transceiver.tx_on = False

    def s1_s10_calibration(self):
        self.test_results.add_line('---------- S1 & S10 Calibration ----------')
        self.radio_tx_off()
        self.radio.set_frequency(Constants.single_align_frequency)
        self.rf_transceiver.tx_freq = Constants.single_align_frequency
        self.rf_transceiver.mod_freq = 0  # no modulation, carrier only
        self.rf_transceiver.tx_on = True  # set tx_on to update tx signal
        # S1 calibration
        power = self.rf_transceiver.set_tx_power(Constants.s1_power_level)
        rssis = np.empty(5, dtype='uint8')
        for i in range(len(rssis)):
            time.sleep(0.1)
            rssis[i] = self.radio.rssi_adc()
        s1_cal = (np.sum(rssis) - np.min(rssis) - np.max(rssis)) / (len(rssis) - 2)
        self.assert_within_range('s1_cal', s1_cal, Constants.sanity_s1_min, Constants.sanity_s1_max)
        if not self.test_results.passed:
            return
        self.radio.write_eeprom_byte(self.radio.gme2.EE_S1_CAL, s1_cal)
        self.test_results.add_line('s1_power = %d dBm, s1_cal = %d' % (power, s1_cal))
        # S10 calibration
        power = self.rf_transceiver.set_tx_power(Constants.s10_power_level)
        for i in range(len(rssis)):
            time.sleep(0.1)
            rssis[i] = self.radio.rssi_adc()
        self.rf_transceiver.tx_on = False
        s10_cal = (np.sum(rssis) - np.min(rssis) - np.max(rssis)) / (len(rssis) - 2)
        self.assert_within_range('s10_cal', s10_cal, Constants.sanity_s10_min, Constants.sanity_s10_max)
        if not self.test_results.passed:
            return
        self.radio.write_eeprom_byte(self.radio.gme2.EE_S10_CAL, s10_cal)
        self.test_results.add_line('s10_power = %d dBm, s10_cal = %d' % (power, s10_cal))

    def noise_level_calibration(self):
        self.test_results.add_line('---------- Noise Level Calibration ----------')
        self.radio_tx_off()
        self.radio.set_frequency(Constants.noise_cal_frequency)
        # Wide-band
        self.radio.set_bandwidth(True)
        adcs = np.empty(10, dtype='uint8')
        for i in range(len(adcs)):
            time.sleep(0.1)
            adcs[i] = self.radio.noise_adc()
        noise_adc = np.average(adcs)
        noise_scale = int(Constants.noise_cal_target * 256 / noise_adc)
        self.assert_within_range('wb_noise_scale', noise_scale, Constants.noise_scale_min, Constants.noise_scale_max)
        if not self.test_results.passed:
            return
        self.radio.write_eeprom_word(self.radio.gme2.EE_NV_WB_SCALE, noise_scale)
        self.test_results.add_line('wb_noise_scale = %d' % noise_scale)
        # Narrow-band
        self.radio.set_bandwidth(False)
        for i in range(len(adcs)):
            time.sleep(0.1)
            adcs[i] = self.radio.noise_adc()
        noise_adc = np.average(adcs)
        noise_scale = int(Constants.noise_cal_target * 256 / noise_adc)
        self.assert_within_range('nb_noise_scale', noise_scale, Constants.noise_scale_min, Constants.noise_scale_max)
        if not self.test_results.passed:
            return
        self.radio.write_eeprom_word(self.radio.gme2.EE_NV_NB_SCALE, noise_scale)
        self.test_results.add_line('nb_noise_scale = %d' % noise_scale)

    def rx_sinad_test(self):
        pass  # ultimate SINAD test using audio output from speaker pins to BeagleBone audio codec

    def tx_sinad_test(self):
        pass  # reverse to the above

    def bluetooth_rf_test(self):
        pass

    def halt_on_error(self, msg=None):
        if msg:
            self.gui.text(text_fg_red + msg + text_reset)
        self.gui.text('Exiting CP50 MTS...')
        try:
            self.power_supply.on = False
        except:
            pass
        try:
            self.rf_transceiver.active = False
        except:
            pass
        exit()

    def print_label(self):
        self.test_results.add_line('---------- Test Done ----------')
        new_sn_date = time.strftime('%y%m')
        if new_sn_date == self.last_sn_date:
            new_sn_count = self.last_sn_count + 1
        else:
            new_sn_count = 1
        new_sn = '%s%s%04d' % (new_sn_date, self.test_station, new_sn_count)
        if use_printer:
            printer.print_label(new_sn, self.test_results.unique_id)
        self.test_results.add_line('serial_no = ' + new_sn)
        self.test_results.serial_no = new_sn
        self.sn_file.write(new_sn + ',' + self.test_results.unique_id + '\n')
        self.sn_file.flush()
        self.last_sn_date = new_sn_date
        self.last_sn_count = new_sn_count

    def finalise_test(self):
        self.rf_transceiver.tx_on = False
        try:
            self.hardware_ptt(False)
            self.audio.stop_play()
        except:
            pass
        try:
            self.radio_tx_off()
            self.radio.write_eeprom_byte(self.radio.gme2.EE_SQL_LEVEL, 2)  # set default squelch level
        except:
            pass
        self.power_supply.on = False
        cx50_manufacturing_test_database.save_results(self.test_results)
        total_time_used = time.time() - self.test_start_time
        self.test_results.add_line('total_time_used = %d seconds' % total_time_used)
        if self.test_results.passed:
            print text_fg_green
            self.test_results.add_line('Test PASSED')
            print text_reset
            self.test_results.save_log()
        else:
            print text_fg_red
            self.test_results.add_line(self.test_results.failure_report_message)
            self.test_results.add_line('Test FAILED')
            print text_reset
            self.test_results.save_log()
        cont = self.gui.question('Continue to next test?')
        if not cont:
            self.com_port.close()
            self.halt_on_error()

    def assert_within_range(self, parameter_name, measured_value, min_value, max_value):
        # This design doesn't allow to have multiple failures in one test. So, check that we're still OK
        if not self.test_results.passed:
            self.halt_on_error('Previous test already failed when testing for %s' % parameter_name)
        self.test_results.__setattr__(parameter_name, measured_value)
        # measured_value is None when the reading fails
        self.test_results.passed = measured_value is not None and (measured_value >= min_value) and (measured_value <= max_value)
        if not self.test_results.passed:
            self.test_results.failure_report_message = '%s out of range (measured: %f, expecting: %f - %f)' % (parameter_name, measured_value, min_value, max_value)

    def assert_true(self, parameter_name, observed_value):
        if not self.test_results.passed:
            self.halt_on_error('Previous test already failed when testing for %s' % parameter_name)
        self.test_results.__setattr__(parameter_name, observed_value)
        # measured_value is None when the reading fails
        self.test_results.passed = (observed_value is not None) and observed_value
        if not self.test_results.passed:
            self.test_results.failure_report_message = '%s is False' % parameter_name


# Initialize the tests
tester = Tester(com_port)

# Run the tests (never returns)
tester.run()
