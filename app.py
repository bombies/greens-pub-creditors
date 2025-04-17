import csv
import datetime
import os
import sys
from multiprocessing import pool
from tkinter import *
from tkinter.filedialog import askopenfile
from tkinter.messagebox import showinfo
from tkinter.ttk import *

import openpyxl

thread_pool = pool.ThreadPool(processes=1)


class Process:
    def __init__(self, root, _file):
        super().__init__()
        self.pb = Progressbar(
            root, orient="horizontal", mode="determinate", length=200
        )
        self.root = root
        self._file = _file

    def __progress(self, amt):
        cur_val = self.pb["value"]
        if cur_val + amt < 100:
            # update progressbar
            self.pb["value"] += amt
            self.root.update_idletasks()
        else:
            # close the progressbar
            self.pb.stop()
            self.pb.destroy()
            showinfo("Success", "The report has been generated successfully.")

    def run(self):
        self.pb.pack()
        self.root.update_idletasks()

        self.__progress(25)

        self.__progress(25)

        crunch_raw_data("./OrderHeaders.csv", "./EmployeeFiles.csv")
        self.__progress(25)

        create_excel_spreadsheet(os.path.dirname(self._file.name))
        self.__progress(25)

    def error(self, err):
        print("Error occurred: {}".format(err))
        self.pb.stop()
        self.pb.destroy()
        showinfo("Error", "An error occurred while processing the file.")

    def start(self):
        print("Starting process...")
        return thread_pool.apply_async(self.run, error_callback=self.error)


def crunch_raw_data(orders_csv_path, employees_csv_path):
    print("Crunching raw data from CSV files...")

    # Creditors will be typed in the form:
    # type Creditor:
    #     Order ID: Int
    #     Creditor Name: String
    #     Amount Credited: Double
    #     Date Credited: Datetime
    #     Cashier: String
    creditorDict = {}

    # Read all open orders from CSV
    all_open_orders = []
    with open(orders_csv_path, "r", newline="") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            # Only include orders with OrderStatus = 1
            if row["OrderStatus"] == "1":
                print(f"Processing order {row['OrderID']}...")
                all_open_orders.append(
                    (
                        row["OrderID"],
                        row["SpecificCustomerName"],
                        float(row["AmountDue"]),
                        row["EmployeeID"],
                        row["OrderDateTime"],
                        row["OrderStatus"],
                    )
                )

    # Read all employees from CSV
    all_employees = []
    with open(employees_csv_path, "r", newline="") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            print(f"Processing employee {row['EmployeeID']}...")
            all_employees.append(
                (row["EmployeeID"], row["FirstName"], row["LastName"])
            )

    # Process the data as before
    for order in all_open_orders:
        print(f"Processing order {order}...")
        order_id = order[0]
        creditor_name = order[1]
        amount_credited = order[2]
        date = order[4]
        employee = next(filter(lambda x: x[0] == order[3], all_employees), None)

        if employee:
            print(f"Found employee {employee[0]} for order {order_id}...")
            cashier = f"{employee[1]} {employee[2]}"
            creditorDict[order_id] = (
                creditor_name,
                amount_credited,
                cashier,
                date,
            )
        else:
            print(
                f"Warning: Employee ID {order[3]} not found for order {order_id}"
            )

    return creditorDict


# This function turns the Creditor abstract type into a
# Excel spreadsheet using openpyxl. The spreadsheet will
# contain the OrderID, Creditor Name, Amount Credited, and
# Cashier as columns.
def create_excel_spreadsheet(path=""):
    print("Crunching raw data...")
    creditorDict = crunch_raw_data("./OrderHeaders.csv", "./EmployeeFiles.csv")

    print("Creating excel spreadsheet...")
    spreadsheet_name = (
        f"{path}/creditors-{datetime.datetime.now().strftime('%d-%m-%Y')}.xlsx"
    )

    # Create workbook if creditors.xlsx does not exist,
    # if it does, the file will just be overwritten.
    try:
        workbook = openpyxl.load_workbook(spreadsheet_name)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        workbook.save(spreadsheet_name)

    worksheet = workbook.active
    cur_row = 1
    print(creditorDict)
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
        _file = askopenfile(
            mode="r", filetypes=[("Microsoft Access Database", "*.mdb")]
        )
        if _file is not None:
            process_thread = Process(root, _file)
            process_thread.start()

    btn = Button(root, text="Open", command=lambda: open_file())
    btn.pack(side=TOP, pady=10)
    root.mainloop()
    return root


if __name__ == "__main__":
    try:
        start_gui()
    except KeyboardInterrupt:
        thread_pool.terminate()
        sys.exit(0)
