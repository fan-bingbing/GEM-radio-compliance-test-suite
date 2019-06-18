import sounddevice as sd
import numpy as np
import matplotlib.pyplot as plt
from scipy.fftpack import fft


fs = 48000 # at least double of signal frequency being sampled
d = 1.5# initial sample duration
gain = 10 # adjustable according to actual volume of received signal

x1 = sd.rec(int(d*fs), dtype='float32', samplerate=fs, channels=1, blocking=True) # record 3 seconds input signal from microphone
x2 = gain*x1[fs:2*fs] # take effective samples from second0.5 to second1 which is 1s duration, fs samples in total
x3 = x2[0:int(10*(fs/1000))] # take first 10ms duration samples
x4 = x3.flatten()# transfer 2D array to 1D, ready for fft operation, fft function can only intake 1D array as input

print(x4) # check x4 content
print(np.size(x4)) # check x4 size
t = 10*np.linspace(0, 1, np.size(x4))# generate 10ms time axis VS samples in first 10ms
print(np.size(t)) # check t size

fig, (ax0, ax1) = plt.subplots(nrows=2)
#ax0.plot(t, x4)
ax0.set_title('Sinusoidal Signal')
ax0.set_xlabel('Time(ms)')
ax0.set_ylabel('Amplitude')
ax0.set_ylim(-2, 2)
line0, = ax0.plot(t, x4)

X = fft(x4)
#generate frequency axis
n = np.size(t)
# frequency axis from 0 to fs/2 (nykuist frequency) which is the first half of fft result
#(2/n) is for magnitude norminalization
fr = (fs/2)*np.linspace(0, 1, int(n/2))
X_m = (2/n)*abs(X[0:np.size(fr)])

ax1.set_title('Magnitude Spectrum')
ax1.set_xlabel('Frequency(Hz)')
ax1.set_ylabel('Magnitude')
ax1.set_ylim(0, 1.5)
line1, = ax1.plot(fr, X_m)
fig.subplots_adjust(hspace=0.6)


while True:
    x1 = sd.rec(int(d*fs), dtype='float32', samplerate=fs, channels=1, blocking=True) # record 3 seconds input signal from microphone
    x2 = gain*x1[fs:2*fs] # take effective samples from second0.5 to second1 which is 1s duration, fs samples in total
    x3 = x2[0:int(10*(fs/1000))] # take first 10ms duration samples
    x4 = x3.flatten()# transfer 2D array to 1D, ready for fft operation, fft function can only intake 1D array as input
    X = fft(x4)
    #generate frequency axis
    n = np.size(t)
    # frequency axis from 0 to fs/2 (nykuist frequency) which is the first half of fft result
    #(2/n) is for magnitude norminalization
    fr = (fs/2)*np.linspace(0, 1, int(n/2))
    # abs is to take magnitude component of fft result array
    # X[0:size(fr)] generate y axis element which has equal size of fr
    X_m = (2/n)*abs(X[0:np.size(fr)])
    line0.set_ydata(x4)
    line1.set_ydata(X_m)
    plt.pause(0.01)# this command will call plt.show()
