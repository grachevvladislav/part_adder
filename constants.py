NEW_HEADER = [
    {'name': 'Комментарий', 'width': 30},
    {'name': 'Количество в ЗИП на 3 год', 'width': 20},
    {'name': 'Количество в ЗИП на 1 год', 'width': 20},
    {'name': 'Количество в оборудовании установлено', 'width': 20},
    {'name': 'Альтернативный партномер', 'width': 20},
    {'name': 'Партномер', 'width': 20},
    {'name': 'Наименование компоненты на русском яз', 'width': 35},
    {'name': 'Описание компоненты на англ яз', 'width': 70},
]

HEADER_WIDTH = [
    NEW_HEADER[0]['width'],
    *[val['width'] for val in NEW_HEADER[3:]],
    20, 10, 25, 20
]
CURRENT_SHEET_NAME = 'Разбивка по серверам'
NEW_SHEET_NAME = 'Сводная таблица'

JSON_FILE_NAME = 'json_conf.xlsx'
