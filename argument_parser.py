import argparse


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
        help='Создание конфигурации из json'
    )
    return parser
