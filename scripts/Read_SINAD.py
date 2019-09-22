import visa
import re
from openpyxl import Workbook

rm = visa.ResourceManager()
print(rm.list_resources())
CMS = rm.open_resource('GPIB0::24::INSTR') # Audio Analyzer
idn_response4 = CMS.query('*IDN?')
print (f"Hello, I am {idn_response4}")


SINAD_file = Workbook()#
ws_SINAD = SINAD_file.active
ws_SINAD.title = "SINAD"

CMS.clear()

for i in range(0,5):
    SINAD_data_str = CMS.query("SINAD:R?")
    SINAD_data_num = re.findall(r'\d+\.\d+', SINAD_data_str)[0]
    print(SINAD_data_num)
    ws_SINAD.cell(row = i+1, column = 1, value = SINAD_data_num)

SINAD_file.save("SINAD.xlsx")
CMS.close()
