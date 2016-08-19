
# öffne die namen, speichere sie alpahbetisch in ein excel dokument

import openpyxl
import re
import codecs

name_regex = r"\s*[A-Za-zäüöÖÜÄ]+\s+" \
             r"[A-Za-zäüöÖÜÄ]+\s+"



name_compiled = re.compile(name_regex, re.VERBOSE)

wb = openpyxl.load_workbook("namelist.xlsx")

sheet = wb.get_sheet_by_name("liste")

f = codecs.open("names.txt", encoding="utf-8")
count = 100
name = f.readlines()
name = "".join(name)
print(name)
print("==========================")
raw_names = name_compiled.findall(name)

raw_names = sorted(raw_names, key=lambda s: s.lower())
print(raw_names)
print("++++++++++++++++++++++++++")
r = 1
c = 1
for n in raw_names:
    cell_ort = sheet.cell(row=r, column=c).coordinate
    print("Hinzugefügt:", n, end="")
    sheet[cell_ort] = n
    r += 1
    if r == 30:
        c += 1
        r = 1


wb.save("namelist.xlsx")
