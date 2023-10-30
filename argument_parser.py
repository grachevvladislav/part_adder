import argparse
from exceptions import SheetNotFound


def configure_argument_parser():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawTextHelpFormatter,
        description=('Калькулятор сводной таблицы ЗИП.\n'
                     'Обязательные столбцы:\n'
                     '|  "E"  |  "F"  |  "G"  | "H"  |"I"|\n'
                     '|EN name|RU name|PN main|PN opt|PCS|'),
        epilog='Без параметров использует первую страницу документа.'
    )
    parser.add_argument(
        'file',
        help='Имя файла *.xls или *.xlsx'
    )
    parser.add_argument(
        '-n',
        '--name',
        help='Выбор листа по имени'
    )
    parser.add_argument(
        '-i',
        '--index',
        help='Выбор листа по номеру, начиная с "0"'
    )
    parser.add_argument(
        '-j',
        '--json',
        help='Создание конфигурации из json',
        action='store_true'
    )
    return parser


def get_sheet(args, file):
    if args.index:
        if args.index in range(len(file.sheetnames)):
            sheet = file.worksheets[int(args.index)]
        else:
            raise SheetNotFound(
                f'В файле {args.file} нет листа номер {args.index}!\n'
                f'Доступны листы от 0 до {len(file.sheetnames) - 1}'
            )
    elif args.name:
        if args.name in file.sheetnames:
            sheet = file[args.name]
        else:
            raise SheetNotFound(
                f'В файле {args.file} нет листа {args.name}!\n'
                f'Доступны: {", ".join(file.sheetnames)}'
            )
    else:
        sheet = file.worksheets[0]
    return sheet
