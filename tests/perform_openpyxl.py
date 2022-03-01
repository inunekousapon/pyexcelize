import resource
from os.path import getsize
from datetime import datetime

from openpyxl import Workbook


def get_maxrss() -> float:
    r = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return r // 1024 // 1024  # bytes on MacOS


print(f"{datetime.now().isoformat()} init: {get_maxrss()}MB")
wb = Workbook()
ws = wb.worksheets[0]
txt = "1234567890" * 10

for row in range(1,50000):
    for col in range(1, 20):
        ws.cell(row=row, column=col, value=txt)

print(f"{datetime.now().isoformat()} writed: {get_maxrss()}MB")
wb.save('./output.xlsx')
print(f"{datetime.now().isoformat()} saved: {get_maxrss()}MB")
wb.close()
print(f"{datetime.now().isoformat()} closed: {get_maxrss()}MB")