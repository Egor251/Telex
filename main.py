import sqlite3
import os
import argparse

import pathlib
from pathlib import Path
path = Path('telex.db')

try:
    import xlrd                                      # Чтение xlsx обязательно версия 1.2.0
except ImportError:
    try:
        os.system('pip3 install xlrd==1.2.0')
        import xlrd
    except Exception:
        os.system('pip install xlrd==1.2.0')
        import xlrd
import inspect
import os
import sys

def get_script_dir(follow_symlinks=True):
    if getattr(sys, 'frozen', False): # py2exe, PyInstaller, cx_Freeze
        path = os.path.abspath(sys.executable)
    else:
        path = inspect.getabsfile(get_script_dir)
    if follow_symlinks:
        path = os.path.realpath(path)
    return os.path.dirname(path)


def create_db(bank):
    conn = sqlite3.connect("telex.db")  # или :memory: чтобы сохранить в RAM
    cursor = conn.cursor()

    # Создание таблиц
    cursor.execute(f'CREATE TABLE IF NOT EXISTS {bank}DD (day int, number int)')
    cursor.execute(f'CREATE TABLE IF NOT EXISTS {bank}CONST (number int)')
    cursor.execute(f"""CREATE TABLE IF NOT EXISTS {bank}MM
                              (month int, number int)
                           """)
    cursor.execute(f"""CREATE TABLE IF NOT EXISTS {bank}SUM
                                  (value int, number int)
                               """)
    cursor.execute(f"""CREATE TABLE IF NOT EXISTS {bank}VARin
                                      (number int, var int)
                                   """)
    cursor.execute(f"""CREATE TABLE IF NOT EXISTS {bank}VARout
                                          (number int, var int)
                                       """)
    cursor.execute(f"""CREATE TABLE IF NOT EXISTS {bank}CUR
                                  (cur text,  number text)""")
    conn.commit()


def select(sql):
    conn = sqlite3.connect("telex.db")
    cursor = conn.cursor()
    cursor.execute(sql)
    return cursor.fetchone()[0]


def db_select(value, base, column,  where, sql=None):
    conn = sqlite3.connect("telex.db")
    cursor = conn.cursor()
    if sql is None:
        if isinstance(where, str):
            cursor.execute(f"SELECT {value} FROM {base} WHERE {column} =  ?;", (where,))
        elif isinstance(where, int):
            cursor.execute(f"SELECT {value} FROM {base} WHERE {column} = {where}")
        else:
            raise ValueError
    else:
        cursor.execute(sql)
    result = cursor.fetchall()
    conn.commit()
    return int(result[0][0])


def update_db(args):

    file = args.file
    variable=args.variable
    bank = args.bank

    def parse_excel(file, variable, bank='EXB'):
        conn = sqlite3.connect("telex.db")
        cursor = conn.cursor()
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        for rownum in range(0, sheet.nrows):
            row = sheet.row_values(rownum)
            for i in range(len(row)):
                try:
                    row[i] = int(row[i])
                except Exception:
                    pass
            cursor.executemany(f"INSERT INTO {bank}{variable} VALUES (?,?)", (row,))
        conn.commit()
    print(f'Таблица {variable} обновлена из файла {file}')

    create_db(bank)
    conn = sqlite3.connect("telex.db")
    cursor = conn.cursor()
    cursor.execute(f'DROP TABLE IF EXISTS {bank}{variable};')
    conn.commit()
    create_db(bank)
    if variable == 'CONST' and isinstance(file, int):
        row = file
        cursor.execute(f"INSERT INTO {bank}{variable} VALUES ({row})")
        conn.commit()
    else:
        parse_excel(file, variable, bank='EXB')


def calc(args):
    message_number = args.message_number
    sum = args.sum
    currency = args.currency
    date = args.date
    operation = args.operation
    bank = args.bank
    var = db_select('var', f'{bank}VAR{operation}', 'number', message_number)
    print(f'VAR = {var}')
    date = date.split('.')
    day = db_select('number', f'{bank}DD', 'day', int(date[0]))
    print(f"Day = {day}")
    month = db_select('number', f'{bank}MM', 'month', int(date[1]))
    print(f"Month = {month}")
    cur = db_select('number', f'{bank}CUR', 'cur', currency)
    print(f"CUR = {cur}")
    sql = f"SELECT * FROM {bank}CONST"
    const = db_select('number', f'{bank}CONST', '1', currency, sql)
    print(f"CONST = {const}")
    print("sum:")
    summary = sum_money(sum)
    print(f"={summary}")
    key = day+month+cur+const+summary+var
    print(f'{day} + {month} + {cur} + {const} + {summary} + {var} = {key}')
    result_key = f'Ключ: {message_number}/{key}'
    print(result_key)
    return result_key


def sum_money(money, bank='EXB'):
    ans = 0
    for x in range(8, -1, -1):
        div = money // (10 ** x)
        if div >= 1:
            if x == 8:
                param = 1
            else:
                param = div
            sql = f'SELECT number FROM {bank}SUM WHERE value = {param*10**x}'
            result = select(sql)
            if x == 8:
                result = result * div
            print(f"    {result}")
            ans += result
            money = money - div*10**x
    return ans


# Объявляем парсер
parser = argparse.ArgumentParser(description='Вычисление ключа',
                                 epilog="Желательно сделать cd в папку со скриптом, но это не точно")
# Подпарсеры
subparsers = parser.add_subparsers(title='subcommands',
                                   description='valid subcommands',
                                   help='Доступные команды. '
                                        'Для большей инфорамции можно посмотреть help по каждой команде')
# Парсер функции update_db
update_parser = subparsers.add_parser('update', help='обновить одну из таблиц в базе')
update_parser.add_argument('--value', dest='variable',
                           help='Таблицы для обновления (SUM, VARin, VARout, CONST, DD, MM, CUR)', required=True)
update_parser.add_argument('--bank', dest='bank', default='EXB',
                           help='Таблицы какого именно банка')
update_parser.add_argument("file", help="Файл с таблицей для обновления "
                                        "(При обновлении CONST вместо файла можно просто вписать число)")
update_parser.set_defaults(func=update_db)

# Парсер для вычисления ключа
key_parser = subparsers.add_parser('key', help='Вычислить ключ')
key_parser.add_argument('-m', dest='message_number',
                        help='Номер сообщения/префикс', required=True, type=int)
key_parser.add_argument('-s', dest='sum',
                        help='Сумма (целое число без знаков препинания и копеек', required=True, type=int)
key_parser.add_argument('-c', dest='currency',
                        help='Валюта (EUR, RUB, UZS или XXX заглавными)', required=True, type=str)
key_parser.add_argument('-d', dest='date',
                        help='Дата перевода в формате DD.MM', required=True, type=str)
key_parser.add_argument('-o', dest='operation',
                        help='Операция (in или out)', required=True, type=str)
key_parser.add_argument('--bank', dest='bank', default='EXB',
                        help='Имя банка (На случай использования разных таблиц)')
key_parser.set_defaults(func=calc)
args = parser.parse_args()


if __name__ == '__main__':

    os.chdir(get_script_dir())

    args = parser.parse_args()
    if not vars(args):
        parser.print_usage()
    else:
        args.func(args)
