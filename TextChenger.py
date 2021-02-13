# from ZWCAD_CONSTANTS import constants

from operator import attrgetter
from win32com import client
import pythoncom
import array
import pythoncom
import re

# input("Нажми Enter для подключения к ZWCAD.Application")
zcad = client.Dispatch("ZWCAD.Application")
doc = zcad.ActiveDocument
model = doc.ModelSpace

# input("Нажми Enter для подключения к Excel.Application")

excel = client.Dispatch("Excel.Application")

#забираем из проекта объект класса SelectionSet
try:
    doc.SelectionSets.Item("SS1").Delete()
except Exception:
    print("Нет набора SS1")
selection_set = doc.SelectionSets.Add("SS1")
print("Выберите элемент на поле Модель и Нажмите Enter")
selection_set.SelectOnScreen()


#из объекта SlelectionSet создаем список элементов и сортируем их по координате X
items_in_selection_set = []
for i in range(selection_set.Count):
    items_in_selection_set.append(selection_set.Item(i))
sorted_selection_set = sorted(items_in_selection_set, key = lambda u: u.InsertionPoint[0] )

sheet = excel.ActiveSheet
selection = excel.Application.Selection
start_row = selection.Cells(1).Row
stop_row = selection.Cells(selection.Cells.Count).Row
column = selection.Cells(1).Column
sheet = excel.ActiveSheet
current_row = start_row

changePatterns = {}

while current_row <= stop_row:
    changePatterns[sheet.Rows(current_row).Cells(column).Value] = sheet.Rows(current_row).Cells(column+1).Value
    current_row = current_row + 1

print(f'Количество замен в файле Excel - {len(changePatterns)}')

prefix = r"\b("
sufix = r"{1})((,)|($){2})"
tempAddSufix = r"\*\(changed\)"

print(f'Принято для замены {selection_set.Count} элемент!')

input("Enter для начала замен")
# for cnt in range(selection_set.Count):
#     item = selection_set.Item(cnt)
#     for key, value in changePatterns.items():
#         pattern = prefix + key + sufix
#         item.TextString = re.sub(pattern, value + "*(changed)" + r'\2', item.TextString)   # r'\2' it is part two of suffix and result value is end of text "$" or ","
#     item.TextString = re.sub(tempAddSufix, "", item.TextString)

for cnt in range(selection_set.Count):
    item = selection_set.Item(cnt)
    result_string = item.TextString
    for key, value in changePatterns.items():
        pattern = prefix + key + sufix
        result_string = re.sub(pattern, value + "*(changed)" + r'\2', result_string)   # r'\2' it is part two of suffix and result value is end of text "$" or ","
    item.TextString = re.sub(tempAddSufix, "", result_string)
