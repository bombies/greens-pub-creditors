import datetime
import os
import sys
import sqlite3
import openpyxl
import pyodbc
import re
from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile
from tkinter.messagebox import showinfo
from collections import namedtuple
from multiprocessing import pool

# Creditors will be typed in the form:
# type Creditor:
#     Order ID: Int
#     Creditor Name: String
#     Amount Credited: Double
#     Date Credited: Datetime
#     Cashier: String
creditorDict = {}

thread_pool = pool.ThreadPool(processes=1)


class Process:
    def __init__(self, root, _file):
        super().__init__()
        self.pb = Progressbar(
            root,
            orient='horizontal',
            mode='determinate',
            length=200
        )
        self.root = root
        self._file = _file

    def __progress(self, amt):
        cur_val = self.pb['value']
        if cur_val + amt < 100:
            # update progressbar
            self.pb['value'] += amt
            self.root.update_idletasks()
        else:
            # close the progressbar
            self.pb.stop()
            self.pb.destroy()
            showinfo("Success", "The report has been generated successfully.")

    def run(self):
        self.pb.pack()
        self.root.update_idletasks()

        mdb_to_sqlite(self._file.name)
        self.__progress(25)

        database = connect_to_database("db_as_sqlite.sqlite")
        self.__progress(25)

        crunch_raw_data(database)
        self.__progress(25)

        create_excel_spreadsheet(os.path.dirname(self._file.name))
        self.__progress(25)

    def start(self):
        return thread_pool.apply_async(self.run)


def mdb_to_sqlite(mdb_file):
    cnxn = pyodbc.connect('Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={};'.format(mdb_file))

    cursor = cnxn.cursor()

    conn = sqlite3.connect("db_as_sqlite.sqlite")
    c = conn.cursor()

    Table = namedtuple('Table', ['cat', 'schem', 'name', 'type'])

    # get a list of tables
    tables = []
    for row in cursor.tables():
        if row.table_type == 'TABLE':
            t = Table(row.table_cat, row.table_schem, row.table_name, row.table_type)
            tables.append(t)

    for t in tables:
        if (t.name != 'OrderHeaders' and t.name != 'EmployeeFiles'):
            continue
        print(t.name)

        # SQLite tables must being with a character or _
        t_name = t.name
        if not re.match('[a-zA-Z]', t.name):
            t_name = '_' + t_name

        # get table definition
        columns = []

        def populate_columns():
            for cursor_row in cursor.columns(table=t.name):
                print('    {} [{}({})]'.format(cursor_row.column_name, cursor_row.type_name, cursor_row.column_size))
                col_name = re.sub('[^a-zA-Z0-9]', '_', cursor_row.column_name)
                optimistic_col = '{} {}({})'.format(col_name, cursor_row.type_name, cursor_row.column_size)
                if optimistic_col not in columns:
                    columns.append(optimistic_col)

        try:
            populate_columns()
        except UnicodeDecodeError:
            def decode_sketchy_utf16(raw_bytes):
                s = raw_bytes.decode("utf-16le", "ignore")
                try:
                    n = s.index('\u0000')
                    s = s[:n]  # respect null terminator
                except ValueError:
                    pass
                return s

            prev_converter = cnxn.get_output_converter(pyodbc.SQL_WVARCHAR)
            cnxn.add_output_converter(pyodbc.SQL_WVARCHAR, decode_sketchy_utf16)
            populate_columns()
            cnxn.add_output_converter(pyodbc.SQL_WVARCHAR, prev_converter)

        cols = ', '.join(columns)

        # create the table in SQLite
        print(f"Creating table {t_name}...\n{cols}")
        c.execute('DROP TABLE IF EXISTS "{}"'.format(t_name))
        print(f'Dropped table {t_name}!')

        try:
            c.execute('CREATE TABLE "{}" ({})'.format(t_name, cols))
            print(f"Created table {t_name}!")
        except sqlite3.OperationalError as e:
            print(f'Couldn\'t create ${t_name}', e)
            continue

        # copy the data from MDB to SQLite
        print(f"Copying data from {t.name} to {t_name}...")
        cursor.execute('SELECT * FROM "{}"'.format(t.name))
        for row in cursor:
            values = []
            for value in row:
                if value is None:
                    values.append(u'NULL')
                else:
                    if isinstance(value, bytearray):
                        value = sqlite3.Binary(value)
                    else:
                        value = u'{}'.format(value)
                    values.append(value)
            v = ', '.join(['?'] * len(values))
            print(f'Appending {values} to {t_name}')
            sql = 'INSERT INTO "{}" VALUES(' + v + ')'
            c.execute(sql.format(t_name), values)

    print("Committing changes to database...")
    conn.commit()
    print("Changes committed!")
    conn.close()


def connect_to_database(database_location):
    return sqlite3.connect(database_location)


def crunch_raw_data(database):
    print("Crunching raw data...")

    # Parse entries into an abstract Creditor type
    database_cursor = database.cursor()
    all_open_orders = database_cursor.execute(
        "SELECT OrderID, SpecificCustomerName, AmountDue, EmployeeID, OrderDateTime, OrderStatus FROM OrderHeaders where OrderStatus = 1"
    ).fetchall()
    all_employees = database_cursor.execute(
        "SELECT EmployeeID, FirstName, LastName FROM EmployeeFiles"
    ).fetchall()
    database_cursor.close()
    database.close()

    db_path = os.path.join(os.getcwd(), "db_as_sqlite.sqlite")
    if os.path.exists(db_path):
        os.remove(db_path)
    else:
        print("The database file didn't exist at {}.".format(db_path))

    for order in all_open_orders:
        order_id = order[0]
        creditor_name = order[1]
        amount_credited = order[2]
        date = order[4]
        employee = list(filter(lambda x: x[0] == order[3], all_employees))[0]
        cashier = f'{employee[1]} {employee[2]}'
        creditorDict[order_id] = (creditor_name, amount_credited, cashier, date)


# This function turns the Creditor abstract type into a
# Excel spreadsheet using openpyxl. The spreadsheet will
# contain the OrderID, Creditor Name, Amount Credited, and
# Cashier as columns.
def create_excel_spreadsheet(path=""):
    print("Creating excel spreadsheet...")
    spreadsheet_name = f"{path}/creditors-{datetime.datetime.now().strftime('%d-%m-%Y')}.xlsx"

    # Create workbook if creditors.xlsx does not exist,
    # if it does, the file will just be overwritten.
    try:
        workbook = openpyxl.load_workbook(spreadsheet_name)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        workbook.save(spreadsheet_name)

    worksheet = workbook.active
    cur_row = 1
    for order_id, order_info in creditorDict.items():
        print(f"Appending order {order_id} to spreadsheet...")

        # Use the order_id and order_info to populate the
        # spreadsheet rows.
        worksheet[f"A{cur_row}"] = order_id
        worksheet[f"B{cur_row}"] = order_info[0]
        worksheet[f"C{cur_row}"] = order_info[1]
        worksheet[f"D{cur_row}"] = order_info[2]
        worksheet[f"E{cur_row}"] = order_info[3]
        cur_row += 1
    workbook.save(spreadsheet_name)
    print(f"Spreadsheet created and saved as {spreadsheet_name}.")


def start_gui():
    root = Tk()
    root.title("Quick Open Order Report")
    root.minsize(300, 100)
    root.geometry("300x100+50+50")

    Label(root, text="Select the database location").pack()

    def open_file():
        _file = askopenfile(mode='r', filetypes=[('Microsoft Access Database', '*.mdb')])
        if _file is not None:
            process_thread = Process(root, _file)
            process_thread.start()

    btn = Button(root, text='Open', command=lambda: open_file())
    btn.pack(side=TOP, pady=10)
    root.mainloop()
    return root


if __name__ == "__main__":
    try:
        start_gui()
    except KeyboardInterrupt:
        thread_pool.terminate()
        sys.exit(0)
