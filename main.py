import os
import sqlite3
from openpyxl import load_workbook


def test():
    print("Hello, World!")


def create_database():
    # データベースをクリアする
    os.remove('sqlite/database.db')
    # データベースを作成する
    db = sqlite3.connect('sqlite/database.db')
    handler = db.cursor()
    # テーブルを作成する
    handler.execute(
        'CREATE TABLE vocabulary(id INTEGER PRIMARY KEY AUTOINCREMENT, word STRING, meaning STRING)')
    # 変更を反映する
    db.commit()
    db.close()


def import_data():
    # ファイルを読み込む
    if os.path.exists('input/kahoot.xlsx'):
        # Kahoot!
        wb = load_workbook('input/kahoot.xlsx')
    elif os.path.exists('input/quizizz.xlsx'):
        # Quizizz
        wb = load_workbook('input/quizizz.xlsx')
    else:
        # File Not Found
        print('Oops!  File "kahoot.xlsx" or "quizizz.xlsx" is not found in "input/".')
        exit(1)
    # データを取り込む
    sheet1 = wb.worksheets[0]
    print(sheet1['A1'].value)


def generate_upload_file():
    if os.path.exists('input/kahoot.xlsx'):
        # Kahoot!
        os.remove('output/kahoot.xlsx')
        generate_kahoot()
        os.remove('input/kahoot.xlsx')
    elif os.path.exists('input/quizizz.xlsx'):
        # Quizizz
        os.remove('output/quizizz.xlsx')
        generate_quizizz()
        os.remove('input/quizizz.xlsx')
    else:
        # 何もしない
        pass


def generate_kahoot():
    # Kahoot!
    wb = load_workbook('template/kahoot.xlsx')
    sheet1 = wb.worksheets[0]
    curRow = 9
    # フォーマット定義
    colQuestion = 'B'
    colOption1 = 'C'
    colOption2 = 'D'
    colOption3 = 'E'
    colOption4 = 'F'
    colTimeLimit = 'G'
    colAnswer = 'H'
    # TODO:
    wb.save('output/kahoot.xlsx')
    pass


def generate_quizizz():
    # Quizizz
    wb = load_workbook('template/quizizz.xlsx')
    sheet1 = wb.worksheets[0]
    curRow = 3
    # フォーマット定義
    colQuestion = 'A'
    colType = 'B'
    colOption1 = 'C'
    colOption2 = 'D'
    colOption3 = 'E'
    colOption4 = 'F'
    colOption5 = 'G'
    colAnswer = 'H'
    colTimeLimit = 'I'
    colImageLink = 'J'
    # TODO:
    wb.save('output/quizizz.xlsx')
    pass


if __name__ == '__main__':
    create_database()
    import_data()
    generate_upload_file()
