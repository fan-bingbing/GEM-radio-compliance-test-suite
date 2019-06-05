'''
Audio test API's for MTS

Documentation: http://wiki.gme.net.au/display/PDLC/CX50+Manufacturing+Software

Copyright (C) 2019 Standard Communications Pty Ltd (GME). All rights reserved.
'''

import sounddevice as sd
import numpy as np
import time
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from matplotlib.widgets import Slider
import scipy.signal as signal
import platform


class SoundCard(object):
    def __init__(self):
        self.samp_rate = 44100
        self.device = ('Microphone Array (Realtek High , MME', 'Speaker/HP (Realtek High Defini, MME')
        sd.query_devices(self.device[0])
        sd.default.device = self.device
        sd.default.channels = 1
        sd.default.samplerate = self.samp_rate

    def play_tone(self, freq, gain=1):
        # generate 1 second of audio data
        t = np.linspace(0, 1.0, self.samp_rate, endpoint=False)
        data = np.sin(2 * np.pi * freq * t) * gain
        # play audio non-blocking continuously
        sd.play(data, loop=True)

    def stop_play(self):
        sd.stop()

    def get_mic_samps(self, num_samps, gain=1):
        samps = sd.rec(num_samps, blocking=True) * gain
        return samps[:, 0]

if __name__ == '__main__':
    audio = SoundCard()



    def update_plot(frame):
        samps = audio.get_mic_samps(10000)
        plt.clf()
        plt.ylim(-1.0, 1.0)
        plt.plot(samps)

    fig = plt.figure()
    ani = animation.FuncAnimation(fig, update_plot, interval=200)
    plt.show()
