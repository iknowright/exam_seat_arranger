from openpyxl import Workbook, load_workbook
from random import shuffle
import argparse
import shutil


AVAILABLE_COLS_IN_EACH_ROOM = {
    '65105': ['B', 'E', 'H', 'K', 'N', 'Q'],
    '4264': ['B', 'D', 'F', 'H', 'J', 'L', 'N'],
    'A1302': ['B', 'D', 'F', 'H', 'J', 'L', 'N', 'Q', 'S']
}
START_ROW = 5
END_ROW = 30

def arrange_seat(student_file, format_file, number):
    arranged_wb = load_workbook(filename=format_file)
    student_wb = load_workbook(filename=student_file)
    student_ws = student_wb.active
    index_list = list(range(1, 2 + int(number)))
    available_rows = range(START_ROW, END_ROW)
    shuffle(index_list)
    counter = 0
    try:
        for room, cols in AVAILABLE_COLS_IN_EACH_ROOM.items():
            ws = arranged_wb[room]
            print(room)
            for col in cols:
                for row in available_rows:
                    if counter > int(number):
                        break
                    if ws[f"{col}{row}"].value == 'x':
                        ws[f"{col}{row}"] = student_ws[f"A{index_list[counter]}"].value
                        counter += 1
    except IndexError:
        pass
    arranged_wb.save(filename=f"arranged_{format_file}")
    arranged_wb.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-s', '--student_file')
    parser.add_argument('-n', '--number')
    parser.add_argument('-f', '--format_file', default='seats.xlsx')
    args = parser.parse_args()

    arrange_seat(args.student_file, args.format_file, args.number)