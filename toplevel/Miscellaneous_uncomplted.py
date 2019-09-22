import sys
import math


# below codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\midlevel')
sys.path.insert(0, r'C:\Users\afan\documents\commercial-radio-test-suite\bottomlevel')
# above codes are for importing files from different directory
# a __init__.py file has to be exist in both "import from" and "import to" folder

import CM60_Test


def CM60():
    Start_frequency = input("Input start frequency in Mhz or press CTRL+C to quit > ")
    Stot_frequency = input("Input stop frequency in Mhz or press CTRL+C to quit > ")
    CM60_Test.ACPR()
