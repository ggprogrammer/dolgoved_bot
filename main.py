import requests as req
from bs4 import BeautifulSoup
import openpyxl
from copy import copy

url = "https://lk.mirea.ru/auth.php"

LOGIN = "sysoev.n.m@edu.mirea.ru"
PASSWORD = "#24680aaA"


# def auth():
#     session = req.Session()
#     log_params = {
#         "AUTH_FORM": "Y",
#         "TYPE": "AUTH",
#         "USER_LOGIN": LOGIN,
#         "USER_PASSWORD": PASSWORD,
#         "USER_REMEMBER": "Y"
#     }

#     r = session.post(url, log_params)
#     r = session.get("https://lk.mirea.ru/learning/debt/")
#     return r.text

# with open("index.html", "w", encoding="utf-8") as f:
#     f.write(auth())

def parser():
    with open("index.html", "r", encoding="utf-8") as f:
        responce = f
        soup = BeautifulSoup(responce, 'lxml')
        block = [i.text for i in soup.find_all("td")[::2]]
        res = []
        for el in block:
            d_name = el[:el.find("(")].strip()
            d_type = el[el.find("(")+1:].split(",")[0].strip()
            d_part = el[el.find("(")+1:].split(",")[1].strip()
            res.append([d_name, d_type, d_part])
        return res
    
def dop_ved_finder(dolgy):
    wb = openpyxl.load_workbook("README.xlsx")
    dop_ved_dict = {}
    for sheet in wb.worksheets:
        if sheet.title == "Справка":
            continue
        empty_counter = 0
        i = 0
        while empty_counter != 5:
            i += 1
            if sheet[f"A{i}"].value is None:
                empty_counter += 1
            elif sheet[f"A{i}"].value in dolgy:
                el = sheet[f"A{i}"].value
                dop_ved_dict.setdefault(sheet.title, []).append([el[:el.find("(")].strip(), el[el.find("("):].strip(), sheet[f"B{i}"].value.capitalize()])
    return dop_ved_dict

def main_info_finder(dolgy):
    wb = openpyxl.load_workbook("README.xlsx")
    main_info_list = []
    for kaf, disc_infos in dolgy.items():
        sheet = wb[kaf]
        for j in range(len(disc_infos)):
            disc = disc_infos[j]
            empty_counter = 0
            print(disc[0])
            res_list = []
            for i in range(1, 100):
                if sheet[f"E{i}"].value is None:
                    empty_counter += 1
                elif disc[0] in sheet[f"E{i}"].value:
                    if [sheet[f"{l}{i}"].value for l in "DEFGHIJKLMN"] not in res_list:
                        res_list.append(i)
            dolgy[kaf][j].append(res_list)
    return dolgy
                    
                
def excel_creator(dop_ved_dict):
    readme_book = openpyxl.load_workbook("README.xlsx")

    book = openpyxl.Workbook()

    sheet = book.active

    sheet["A1"] = "Кафедра"
    sheet["B1"] = "Дисциплина"
    sheet["C1"] = "Ведомость/допуск"



    i = 1
    for kaf, disc_info in dop_ved_dict.items():
        i += 1
        sheet[f"A{i}"] = kaf
        readme_sheet = readme_book[kaf]
        for disc in disc_info:
            sheet[f"B{i}"] = disc[0] + " " + disc[1]
            sheet[f"C{i}"] = disc[2]
            for el in disc[3]:
                l_str = "EFGHIJKLMNO"
                readme_str = "DEFGHIJKLMN"
                for j in range(len(l_str)):
                    sheet[f"{l_str[j]}{i}"] = readme_sheet[f"{readme_str[j]}{el}"].value
                    sheet[f"{l_str[j]}{i}"].font = copy(readme_sheet[f"{readme_str[j]}{el}"].font)
                    sheet[f"{l_str[j]}{i}"].alignment = copy(readme_sheet[f"{readme_str[j]}{el}"].alignment)
                    sheet[f"{l_str[j]}{i}"].border = copy(readme_sheet[f"{readme_str[j]}{el}"].border)
                    sheet[f"{l_str[j]}{i}"].fill = copy(readme_sheet[f"{readme_str[j]}{el}"].fill)

                    sheet[f"{l_str[j]}{i}"].number_format = readme_sheet[f"{readme_str[j]}{el}"].number_format
                i += 1
            i += 1

    book.save("results.xlsx")
    book.close()


def printer(dop_ved_dict):
    for kaf, disc_info in dop_ved_dict.items():
        print(kaf) 
        for disc in disc_info:
            print(disc[0] + " " + disc[1])
            print(disc[2])
            for item in disc[3]:
                print(item)
            print("\n")


if __name__ == "__main__":
    dolgy = parser()
    dolgy = [f'{i[0]} ({i[1]})' for i in dolgy]
    dop_ved_dict = dop_ved_finder(dolgy)
    dop_ved_dict = main_info_finder(dop_ved_dict)
    excel_creator(dop_ved_dict)
    