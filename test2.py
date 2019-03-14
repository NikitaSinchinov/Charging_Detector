import matplotlib.pyplot as plt
import numpy as np

x = np.random.rand(15)
y = np.random.rand(15)
names = np.array(list("ABCDEFGHIJKLMNO"))

fig, ax = plt.subplots()


annot = ax.annotate("AZAZA", xy=(0,0), xytext=(20,20), textcoords="offset points")




plt.show()