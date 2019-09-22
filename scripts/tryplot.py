import matplotlib.pyplot as plt
import numpy as np
import time

# a = np.arange(0, 10, 0.2)
# b = np.sin(a)
# c = np.cos(a)
# d = np.random.rand(2)
# e = np.random.randn(3, 4)

# print(a)
# print(b)
# print(c)
# print(d)
# print(e)

# fig, ax = plt.subplots()
# ax.plot(a, b)
# plt.show()



fig, ax = plt.subplots()
line, = ax.plot(np.random.randn(100))

tstart = time.time()
num_plots = 0
while time.time()-tstart < 1:
    line.set_ydata(np.random.randn(100))
    plt.pause(0.001)
    # fig.canvas.draw()
    # fig.canvas.flush_events()
    num_plots += 1
print(num_plots)





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
