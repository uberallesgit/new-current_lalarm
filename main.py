
import openpyxl
import json
from datetime import datetime
current_time = datetime.now().strftime('%d_%m_%H_%M_%S')
print(current_time)

our_book_path = "БС Феодосия.xlsx"
current_alarms_path = "CurrentAlarmsTest.xlsx"
address_list_path = "БС Феодосия.xlsx"

# Читаем наши станции
def read_our_book(path):
    book = openpyxl.load_workbook(filename=path)
    # print(book.worksheets)
    our_bs_list = []
    our_addr_list = []
    counter = 0
    for sheet in book.worksheets:
        for i in range(1, 150):
            if sheet[f'a{i}'].value != None:
                counter += 1
                our_bs_list.append(sheet[f'a{i}'].value)
                our_addr_list.append(sheet[f'b{i}'].value)

    return our_bs_list, our_addr_list

#Читаем Current_alarms
def read_current_alarms(path,bs_list, our_addr_list):
    book = openpyxl.load_workbook(filename=path)
    our_al_list = []
    minor_al_list = []
    minor_count = 0
    major_count = 0
    # col = input("Введите букву колонки:")
    # if col == "":
    col = "g"
    for sheet in book.worksheets:
        for i in range(1, 4000):
            if sheet[f'{col}{i}'].value in bs_list:
                if sheet[f'{col}{i}'].value != None:
                    if sheet[f'b{i}'].value == "Minor":
                        minor_count +=1
                        minor_al_list.append(
                            {"BS_name": sheet[f'{col}{i}'].value,
                             "Severity": sheet[f'b{i}'].value,
                             "Alarm_name": sheet[f'd{i}'].value,
                             "Location Information":sheet[f'h{i}'].value,
                             "NE Type": sheet[f'e{i}'].value,
                             "First-occured": sheet[f'j{i}'].value,
                             "Last_occured": sheet[f'k{i}'].value,
                             "number": minor_count,
                             }
                        )
                    else:
                        major_count +=1
                        our_al_list.append(
                            {"BS_name": sheet[f'{col}{i}'].value,
                             "Severity": sheet[f'b{i}'].value,
                             "Alarm_name": sheet[f'd{i}'].value,
                             "Location Information": sheet[f'h{i}'].value,
                             "NE Type": sheet[f'e{i}'].value,
                             "First-occured": sheet[f'j{i}'].value,
                             "Last_occured": sheet[f'k{i}'].value,
                             "Number": major_count
                             }
                        )
    # print("Наши  алармы :",our_al_list)
    # print(len(our_al_list))
    return our_al_list, minor_al_list,minor_count,major_count


list_1,list_4 = read_our_book(path=our_book_path)
# list_4 = address_dict(address_list_path)
list_2,list_3,minor_count, major_count = read_current_alarms(path=current_alarms_path,bs_list=list_1,our_addr_list=list_4)


with open(f"alarms/{major_count}_maj_{current_time}.json","w") as file:
    json.dump(list_2,file, indent=4, ensure_ascii=False)

with open(f"alarms/{minor_count}_min_{current_time}.json","w") as file:
    json.dump(list_3,file, indent=4, ensure_ascii=False)

