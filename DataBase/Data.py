import sqlite3
from pathlib import Path

# Database class.
class AddressersData:
    def __init__(self, database, table):
        self.db = database
        self.table = table

        # Establish connection with the Database.
        self.conn = sqlite3.connect(database)
        self.cursor = self.conn.cursor()

    def insert(self, name, email, keyword):
        """
        Insert addressee data into Database.
        """

        insert = f'INSERT OR IGNORE INTO  {self.table}' \
            '(name, email, keyword) Values (?, ?, ?)'
        self.cursor.execute(insert, (name, email, keyword))
        self.conn.commit()

    def edit(self, nome, email, keyword, id):  # Edit addressee data by ID.
        edit = f'UPDATE {self.table} SET name=?, email=?, keyword=? WHERE id=?'
        self.cursor.execute(edit, (nome, email, keyword, id))
        self.conn.commit()

    def delete(self, id):
        """
        Remove addressee from Database by ID.
        """

        delete = f'DELETE FROM {self.table} WHERE id=?'
        self.cursor.execute(delete, (id,))
        self.conn.commit()

    def people_data(self, kword):
        """
        Return addressee E-mail from Database by his keyword.
        """

        data_list = []
        self.cursor.execute(
            f'SELECT * FROM {self.table}')
        for i in self.cursor.fetchall():
            data_list.append(i)
        for i in data_list:
            if kword in i:
                return i[2]

    def close(self):  # End the connection with Database.
        self.cursor.close()
        self.conn.close()


if __name__ == '__main__':
    addresseers_data = AddressersData(
        Path('DataBase/PeopleData.db'), 'PeopleEmails')

    addresseers_data.close()
