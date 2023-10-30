# part_adder

- Приложение для автоматизации процесса расчёта ЗИП для серверов. Создание 
сводной таблицы, расчёт необходимого для закупки количества деталей в 
зависимости от срока поддержки оборудования (1 и 3 года).
- Создание таблицы спецификации по json. Сервера HP Proliant Gen9/10.

### Запуск

1. Cоздать и активировать виртуальное окружение:
```shell
python -m venv venv
source venv/Scripts/activate
```
2. Установить зависимости из файла requirements.txt:
```shell
pip install -r requirements.txt
```
### I Рассчёт ЗИП
3. Поместить файл `*.xlsx` в папку проекта. Образец файла: `example.xlsx`.
4. Выполнить файл (пример для 2-го листа `example.xlsx`):
    ```shell
    python main.py example.xlsx -i 2
    ```
#### Обязательные столбцы:

|  "E"  |  "F"  |  "G"  | "H"  | "I" |
| ----- | ----- | ----- | ---- |-----|
|EN name|RU name|PN main|PN opt| PCS |

#### Аргументы:

**file** - Имя файла *.xls или *.xlsx;\
**'-n', '--name'** - выбор листа по имени;\
**'-i', '--index'** - выбор листа по номеру, начиная с "0";\

Без аргументов используется первая страница документа (index=0). В результате 
выполнения программы создается лист `"Сводная таблица"` содержащий перечень 
компонент, установленных во всех серверах и необходимое для закупки их 
количество.

### II Создание спецификации
3. Поместить файл в кодировке `utf-8` в папку проекта. 
Образец файла: `example_json.txt`.
4. Выполнить файл:
    ```shell
    python main.py example_json.txt -j
    ```
Результатом работы является файл `json_conf.xlsx` с информацией о серверах, 
собранной и файла.

#### Аргументы:

**file** - Имя файла в формате utf-8;\
**'-j', '--json'** - флаг для входа в режим парсинга json;\
**'-f', '--force'** - Принудительное создание файла Exel при парсинге json;\