
import numpy as np
import matplotlib.pyplot as plt
from scipy.fftpack import fft

fs = 1000 # at least double of signal frequency being sampled
t = np.arange(0, 1, 1/fs)
f = 20 # human ear response from 20-20kHz

x = np.sin(2*np.pi*f*t) + 0.5*np.sin(2*np.pi*40*t) + 1.5*np.sin(2*np.pi*5*t) # 3 freq sinwave combination

plt.subplot(2,1,1)
plt.plot(t, x);plt.title('Sinusoidal Signal')
plt.xlabel('Time(s)');plt.ylabel('Amplitude')
#plt.show()
#generate frequency axis
n = np.size(t)
# frequency axis from 0 to fs/2 (nykuist frequency) which is the first half of fft result
fr = (fs/2)*np.linspace(0, 1, n/2)

X = fft(x)
X_m = (2/n)*abs(X[0:np.size(fr)])#(2/n) is for magnitude normilization




plt.subplot(2,1,2)
plt.plot(fr, X_m)
plt.show()
