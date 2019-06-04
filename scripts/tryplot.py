import matplotlib.pyplot as plt
import numpy as np

a = np.arange(0, 10, 0.2)
b = np.sin(a)
c = np.cos(a)
# d = np.random.rand(2, 3)
# e = np.random.randn(3, 4)

# print(a)
# print(b)
# print(c)
# print(d)
# print(e)

fig, ax = plt.subplots()
ax.plot(a, b, c)
plt.show()
