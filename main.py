import ast
import os.path
from code.argument_parser import configure_argument_parser, get_sheet
from code.calculator import zip_calculation
from code.constants import (
    CURRENT_SHEET_NAME,
    HEADER_WIDTH,
    JSON_FILE_NAME,
    PIVOT_HEADER,
    NEW_SHEET_NAME,
)
from code.exceptions import FileAlreadyExists, InvalidFileFormat, SheetNotFound
from code.json_parser import get_json_data
from code.exel import create_main_table
from collections import namedtuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.exceptions import InvalidFileException


def get_data(sheet):
    data_counter = {}
    data_info = {}
    Part = namedtuple("Part", "en_name ru_name alternative_pn comment")
    sheet.title = CURRENT_SHEET_NAME
    # fix style
    sheet.row_dimensions[1].height = 55
    sheet.sheet_view.zoomScale = 55
    for col in sheet.iter_cols(max_col=len(PIVOT_HEADER), max_row=1):
        sheet.column_dimensions[
            col[0].column_letter
        ].width = HEADER_WIDTH.pop()
        col[0].alignment = Alignment(
            wrap_text=True,
            vertical="top",
        )
        col[0].font = Font(
            name="Helvetica Neue (Headings)",
            size=13,
            bold=True,
            color="FF000000",
        )
        col[0].border = Border(
            left=Side(border_style="thin", color="FF000000"),
            right=Side(border_style="thin", color="FF000000"),
            top=Side(border_style="thin", color="FF000000"),
            bottom=Side(border_style="thin", color="FF000000"),
        )

    for string in range(2, sheet.max_row):
        if sheet.cell(string, 9).value == "":
            for column in range(1, sheet.max_column):
                cell = sheet.cell(string, column)
                if cell.value is not None:
                    print(f"Ошибка количества в строке номер {string}")
                    break
                cell.font = Font(
                    name="Helvetica Neue (Headings)", size=13, color="FF000000"
                )
                cell.border = Border(
                    left=Side(border_style="thin", color="FF000000"),
                    right=Side(border_style="thin", color="FF000000"),
                    top=Side(border_style="thin", color="FF000000"),
                    bottom=Side(border_style="thin", color="FF000000"),
                )
                cell.fill = PatternFill(
                    fill_type="solid", start_color="DCE2F1"
                )
        try:
            count = sheet.cell(string, 9).value
            if count is None:
                continue
            count = int(count)
        except ValueError:
            print(
                f'"{sheet.cell(string, 9).value}" - '
                f"недопустимое значение ячейки "
                f"{sheet.cell(string, 9).coordinate}"
            )
            break
        en_name = sheet.cell(string, 5).value
        ru_name = sheet.cell(string, 6).value
        pn = sheet.cell(string, 7).value
        alternative_pn = sheet.cell(string, 8).value
        comment = sheet.cell(string, 10).value
        data_counter[pn] = data_counter.get(pn, 0) + count
        if pn not in data_info.keys():
            data_info[pn] = Part(en_name, ru_name, alternative_pn, comment)
    print(f"Обработано {sheet.max_row} строк!")
    return data_counter, data_info


def create_pivot_table(file, data_counter, data_info, zip_data):
    # create new sheet
    if NEW_SHEET_NAME in file.sheetnames:
        file.remove(file[NEW_SHEET_NAME])
    sheet = file.create_sheet(NEW_SHEET_NAME)
    # sheet settings
    # file.active = sheet
    sheet.sheet_view.zoomScale = 55
    # header
    sheet.row_dimensions[1].height = 55
    for col in sheet.iter_cols(max_col=len(PIVOT_HEADER), max_row=1):
        title = PIVOT_HEADER.pop()
        col[0].value = title["name"]
        sheet.column_dimensions[col[0].column_letter].width = title["width"]
        col[0].alignment = openpyxl.styles.Alignment(
            wrap_text=True,
            vertical="top",
        )
        col[0].font = Font(
            name="Helvetica Neue (Headings)",
            size=13,
            bold=True,
            color="FF000000",
        )
        col[0].border = Border(
            left=Side(border_style="thin", color="FF000000"),
            right=Side(border_style="thin", color="FF000000"),
            top=Side(border_style="thin", color="FF000000"),
            bottom=Side(border_style="thin", color="FF000000"),
        )
    # add data
    for row in sheet.iter_rows(
        max_col=len(PIVOT_HEADER), min_row=2, max_row=len(data_info) + 1
    ):
        current_pn = list(data_info.keys())[row[0].row - 2]
        sheet.row_dimensions[row[0].row].height = 35
        unit = [
            data_info[current_pn].en_name,
            data_info[current_pn].ru_name,
            current_pn,
            data_info[current_pn].alternative_pn,
            data_counter[current_pn],
            *zip_data[current_pn],
            data_info[current_pn].comment,
        ]
        for cell in row:
            cell.value = unit[cell.col_idx - 1]
            cell.alignment = openpyxl.styles.Alignment(
                wrap_text=True,
                vertical="center",
                horizontal="center",
            )
            cell.font = Font(
                name="Helvetica Neue (Headings)", size=13, color="FF000000"
            )
            cell.border = Border(
                left=Side(border_style="thin", color="FF000000"),
                right=Side(border_style="thin", color="FF000000"),
                top=Side(border_style="thin", color="FF000000"),
                bottom=Side(border_style="thin", color="FF000000"),
            )
            cell.fill = PatternFill(fill_type="solid", start_color="DCE2F1")


def main() -> None:
    args = configure_argument_parser().parse_args()
    if args.json:
        try:
            with open(args.file) as file:
                file_contents = file.read()
            if os.path.isfile(JSON_FILE_NAME) and not args.force:
                raise FileAlreadyExists(
                    f"Файл {JSON_FILE_NAME} уже существует. Переименуйте или "
                    f"удалите файл.\nПринудительно создать файл: -f/--force"
                )
        except FileNotFoundError:
            print(f"Файл {args.file} не найден!")
        except UnicodeDecodeError:
            print("Файл не поддерживается")
        except SyntaxError:
            print("Синтаксическая ошибка в файле")
        except FileAlreadyExists as e:
            print(e)
        else:
            file = openpyxl.Workbook()
            sheet = file.worksheets[0]
            servers = get_json_data(ast.literal_eval(file_contents))
            create_main_table(sheet, servers)
            file.save(JSON_FILE_NAME)
    else:
        try:
            file = openpyxl.load_workbook(filename=args.file)
            sheet = get_sheet(file, args)
            if sheet.max_row < 2 or sheet.max_column < 9:
                raise InvalidFileFormat("Неверный формат таблицы!")
        except InvalidFileException:
            print("Файл не поддерживается! Нужен файл .xlsx,.xlsm,.xltx,.xltm")
        except (SheetNotFound, InvalidFileFormat) as e:
            print(e)
        else:
            data_counter, data_info = get_data(sheet)
            zip_data = zip_calculation(data_counter, data_info)
            create_pivot_table(file, data_counter, data_info, zip_data)
            file.save(args.file)
    return


if __name__ == "__main__":
    main()
