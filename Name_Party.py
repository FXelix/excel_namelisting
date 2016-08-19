
# öffne die namen, speichere sie alpahbetisch in ein excel dokument

import openpyxl
import re
import codecs

# regex: forename and surname
name_regex = r"[A-Za-zäüöÖÜÄ]+\s+" \
             r"[A-Za-zäüöÖÜÄ]+\s+"

name_compiled = re.compile(name_regex, re.VERBOSE)

#open the excel document
wb = openpyxl.load_workbook("namelist.xlsx")
sheet = wb.get_sheet_by_name("liste")

# open file with names
f = codecs.open("names.txt", encoding="utf-8")
name = f.readlines()
name = "".join(name)
print(" Raw Names ".center(40,"="), "\n")

#find names that match the pattern
raw_names = name_compiled.findall(name)
raw_names = sorted(raw_names, key=lambda s: s.lower())  #  sort them alphabetically
print(raw_names)
print(" Alle Namen ".center(40,"="))

# info about names and lines
count_names = len(raw_names)
line_count = sum(1 for line in open('names.txt'))
print("{} Namen gefunden bei {} Zeilen!".format(count_names, line_count))
if count_names == line_count:
    print("Volkommene Übereinstimmung!")
else:
    print("Es wurden weniger Namen gefunden als Zeilen im Dokument.")
print("-" * 40)

r = 1   # starting row 1
c = 1   # starting column A

# iterate over all names found and print them into the excel document
for n in raw_names:
    cell_ort = sheet.cell(row=r, column=c).coordinate
    print("Hinzugefügt:", n, end="")
    sheet[cell_ort] = n
    r += 1
    if r == 40:
        c += 1
        r = 1

#  quit and save
wb.save("namelist.xlsx")
f.close()
