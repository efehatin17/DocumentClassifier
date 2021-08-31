from openpyxl import load_workbook
import re
import pandas as pd

def vectorize(sheet, column, number_list):
    big_name_list = []
    encodings_dict = dict()
    for i in range(2, 1042):
        names = sheet[column + str(i)].value
        names = names.replace(",", ", ")
        names = names.split("/")[1]
        x = re.findall("\d+", names)
        if (x):
            total = re.split('\d+', names)[1]
        else:
            continue
        total = total.replace(") ", "")
        name_list = total.split(", ")
        name_list = [name.lower() for name in name_list]
        name_list = [name.lstrip() for name in name_list]
        big_name_list += name_list
        encodings_dict[number_list[i - 2]] = name_list
    final_list = list(dict.fromkeys(big_name_list))
    final_list = sorted(final_list)
    final_list = final_list[1:]
    return final_list, encodings_dict

def find_number(sheet):
    number_list = []
    for i in range(2, 1042):
        filename = sheet["C" + str(i)].value
        filename.replace(" ", "")
        number_of_dashes = filename.count("-")
        if number_of_dashes == 5 or number_of_dashes == 4:
            number = filename.split("-")[4]
            number_list.append(number)
    return number_list

def hot_encode(list, dict):
    encodings_dict = dict
    for key in dict.keys():
        encodings_dict[key] = [1 if i in dict[key] else 0 for i in list]
    return encodings_dict


if __name__ == "__main__":
    workbook = load_workbook(filename="Lead Monitoring Report  (51) (1).xlsx")
    sheet = workbook.active
    numbers_list = find_number(sheet)

    R_list, R_dict = vectorize(sheet, "Q", numbers_list)
    LR_list, LR_dict = vectorize(sheet, "R", numbers_list)
    OLR_list, OLR_dict = vectorize(sheet, "S", numbers_list)

    encoding = hot_encode(R_list, R_dict)
    print(R_list)
    print(encoding)

    df = pd.read_csv('output.txt', delimiter="\t")






