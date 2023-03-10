import openpyxl
import argparse
from exceptions import NoSheet
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill


HEADER = [
    {'name': 'Комментарий', 'width': 30},
    {'name': 'Количество в ЗИП на 1 год', 'width': 20},
    {'name': 'Количество в оборудовании установлено', 'width': 20},
    {'name': 'Альтернативный партномер', 'width': 20},
    {'name': 'Партномер', 'width': 20},
    {'name': 'Наименование компоненты на русском яз', 'width': 35},
    {'name': 'Описание компоненты на англ яз', 'width': 70},
]
CURRENT_SHEET_NAME = 'Разбивка по серверам'
NEW_SHEET_NAME = 'Сводная таблица'


def configure_argument_parser():
    parser = argparse.ArgumentParser(description='Калькулятор ЗИП')
    parser.add_argument(
        'file',
        help='Имя файла *.xls или *.xlsx'
    )
    parser.add_argument(
        '-n',
        '--name',
        help='Название листа в файле'
    )
    parser.add_argument(
        '-i',
        '--index',
        help='Номер листа в файле'
    )
    return parser


def get_data(sheet):
    data_counter = {}
    data_info = {}
    for string in range(2, sheet.max_row):
        try:
            count = int(sheet.cell(string, 9).value)
        except TypeError:
            for column in range(1, sheet.max_column):
                if sheet.cell(string, column).value is not None:
                    print(f'Ошибка колличества в строке номер {string}')
                    break
            continue
        en_name = sheet.cell(string, 5).value
        ru_name = sheet.cell(string, 6).value
        pn = sheet.cell(string, 7).value
        alternative_pn = sheet.cell(string, 8).value
        data_counter[pn] = data_counter.get(pn, 0) + count
        if pn not in data_info.keys():
            data_info[pn] = [en_name, ru_name, alternative_pn]
    return data_counter, data_info


def edit_sheet(sheet):
    sheet.title = CURRENT_SHEET_NAME


def create_new_sheet(file, data_counter, data_info):
    # create new sheet
    if NEW_SHEET_NAME in file.sheetnames:
        file.remove(file[NEW_SHEET_NAME])
    sheet = file.create_sheet(NEW_SHEET_NAME)
    sheet.row_dimensions[1].height = 55
    for col in sheet.iter_cols(max_col=len(HEADER), max_row=1):
        title = HEADER.pop()
        col[0].value = title['name']
        sheet.column_dimensions[col[0].column_letter].width = title['width']
        col[0].alignment = openpyxl.styles.Alignment(
            wrap_text=True,
            vertical='top',
        )
        col[0].font = Font(
            name='Helvetica Neue (Headings)',
            size=13,
            bold=True,
            color='FF000000'
        )
        col[0].border = Border(
            left=Side(border_style='thin', color='FF000000'),
            right=Side(border_style='thin', color='FF000000'),
            top=Side(border_style='thin', color='FF000000'),
            bottom=Side(border_style='thin', color='FF000000'),
        )
    # add data
    for row in sheet.iter_rows(
            max_col=len(HEADER),
            min_row=2,
            max_row=len(data_info) + 1
    ):
        current_pn = list(data_info.keys())[row[0].row - 2]
        sheet.row_dimensions[row[0].row].height = 35
        unit = [
            *data_info[current_pn][:-1], current_pn,
            data_info[current_pn][-1], data_counter[current_pn],
            '', ''
        ]
        for cell in row:
            cell.value = unit[cell.col_idx - 1]
            cell.alignment = openpyxl.styles.Alignment(
                wrap_text=True,
                vertical='center',
                horizontal='center',
            )
            cell.font = Font(
                name='Helvetica Neue (Headings)',
                size=13,
                color='FF000000'
            )
            cell.border = Border(
                left=Side(border_style='thin', color='FF000000'),
                right=Side(border_style='thin', color='FF000000'),
                top=Side(border_style='thin', color='FF000000'),
                bottom=Side(border_style='thin', color='FF000000'),
            )
            cell.fill = PatternFill(
                fill_type='solid',
                start_color='DCE2F1'
            )


def main():
    args = configure_argument_parser().parse_args()
    try:
        file = openpyxl.load_workbook(filename=args.file)
        if args.index is not None:
            sheet = file.worksheets[int(args.index)]
        elif args.name is not None:
            sheet = file[args.name]
        else:
            raise NoSheet
        edit_sheet(sheet)
        data_counter, data_info = get_data(sheet)
        create_new_sheet(file, data_counter, data_info)

        file.save(filename=args.file)
    except FileNotFoundError:
        print(f'Файл {args.file} не найден!')
    except NoSheet:
        print('Лист не выбран!')
    except IndexError:
        print(f'В файле {args.file} нет такого листа')
    if sheet.max_row < 2 and sheet.max_column < 9:
        print('Неверный формат таблицы!')
        return


if __name__ == '__main__':
    main()
