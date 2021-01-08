import os
import sqlite3


def test():
    print("Hello, World!")


def create_database():
    os.remove('sqlite/database.db')
    db = sqlite3.connect('sqlite/database.db')
    handler = db.cursor()
    handler.execute(
        'CREATE TABLE vocabulary(id INTEGER PRIMARY KEY AUTOINCREMENT, word STRING, meaning STRING)')
    db.commit()
    db.close()


if __name__ == '__main__':
    create_database()
    test()
