import os
import typing

import xlrd2

DirName = str
Intrest = float
Detail_Intrest_Expla = list[tuple[Intrest, str]]

def main():
    data: list[tuple[DirName, Intrest, Detail_Intrest_Expla]] = read_xls_files()

    with open("справка.txt", 'w', encoding='utf-8') as f:
        for dat in data:
            f.write(f"{dat[1]} -- {dat[0]}\n")

    with open("справка - подробна.txt", 'w', encoding='utf-8') as f:
        for dat in data:
            f.write(f"{dat[1]} -- {dat[0]}\n")

            for detail in dat[2]:
                f.write(f"\t\t{detail[0]} = {detail[1]}\n")
            f.write("\n")



def read_xls_files()-> list[tuple[DirName, Intrest, Detail_Intrest_Expla]]:
    xls_files = get_list_of_files(".")

    xls_info: list[tuple[DirName, Intrest, Detail_Intrest_Expla]] = []

    for dirr in xls_files:

        individual_detailed: Detail_Intrest_Expla = []
        interest_rate = 0.0
        for file_name in dirr.files:
            # book = xlrd2.open_workbook("C:\\Users\\j1ko\Music\\Desktop\\2023 фев\\СИМОНА - ЯНКО СИМЕОНОВ\\payments_3_27-02-2023.xls")
            book = xlrd2.open_workbook(f"{dirr.folder_name}\\{file_name}")
            # print("The number of worksheets is {0}".format(book.nsheets))
            # print("Worksheet name(s): {0}".format(book.sheet_names()))

            sh = book.sheet_by_index(0)
            if type(sh) == list:
                raise "Big Problem"

            for rx in range(sh.nrows):
                row_line = sh.row(rx)

                type_obligation: str = row_line[4].value

                has_lihvi = ("лихва" in type_obligation) or ("лихви" in type_obligation)
                if not has_lihvi:
                    continue

                type_document = row_line[5].value
                number_document = row_line[6].value
                date: str = row_line[7].value
                period: str = row_line[8].value
                srok: str = row_line[9].value


                lihva = float(row_line[2].value)
                individual_detailed.append((lihva, f"{type_obligation} | {type_document} | {number_document} | дата: {date} | период: {period} | {srok}"))
                interest_rate += lihva
                # print(line_str)
                # print(has_lihvi)
                # print(lihva)
                # print(type(lihva))

        # print(dirr.folder_name, " =", end=" ")
        # print(round(interest_rate, 2))
        # print(individual_detailed)
        xls_info.append((dirr.folder_name, round(interest_rate, 2), individual_detailed))

    return xls_info

class DirFiles(typing.NamedTuple):
    folder_name: str
    files: list[str]


def get_list_of_files(dir_name: str) -> list[DirFiles]:
    list_of_file = os.listdir(dir_name)


    sub_dirs: list[DirFiles] = []

    for entry in list_of_file:
        full_path = os.path.join(dir_name, entry)

        if os.path.isdir(full_path):
            list_sub_dir = os.listdir(full_path)
            sub_dirs.append(DirFiles(full_path, [file for file in list_sub_dir]))

    return sub_dirs

if __name__ == '__main__':
    main()