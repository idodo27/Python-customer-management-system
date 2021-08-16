import pyodbc
from tkinter import *
import datetime
import random

"""Authors: Ido Nidbach & Michael Winkler"""

"""Creating the main window :"""

company = Tk()
company.title("Communications LTD")
company.configure(background="sky blue")
company.geometry("640x480")

"""Global variables :"""

user_option = IntVar()
package_id = StringVar()
speed = StringVar()


"""Dictionary used to store the user's input when adding a new customer"""
client_stringVars = {"first_name": StringVar(), "last_name": StringVar(), "birth_date": StringVar(),
                     "city": StringVar(), "state": StringVar(), "street": StringVar(), "main_phone": StringVar(),
                     "sec_phone": StringVar(), "fax": StringVar(), "discount": DoubleVar()}

"""Dictionary used to store the user's input when adding a new package"""
package_stringVars = {"speed": StringVar(), "start_date": StringVar(), "monthly_payment": IntVar(),
                      "sector_id": IntVar()}


"""Database connection details : """

conn = pyodbc.connect('Driver={SQL SERVER};'
                      'Server=DOD\SQLEXPRESS;'
                      'Database=Communications_LTD;'
                      'Trusted_Connection=yes;'
                      'Encrypt=yes;'
                      'TrustServerCertificate=yes;')

"""Button functions : """


def execreate_customer():
    """This function is used to add a new customer to the table via the "submit" button.
    After adding the customer the entry fields will be cleared and the all the customers will be displayed.
    """
    create_customer(conn)
    for key in client_stringVars:
        client_stringVars[key].set("")
    read_customers(conn)


def execreate_package():
    """This function is used to add a new package to the table via the "submit" button.
    After adding the package the entry fields will be cleared and all the packages will be displayed.
    """

    create_package(conn)
    for key in package_stringVars:
        package_stringVars[key].set("")
    read_packages(conn)


def exeupdate_package():
    """This function is used to edit an existing package to the table via the "submit" button.
    After editing the package the entry fields will be cleared abd all the packages will be displayed.
    """

    update_package(conn)
    package_id.set("")
    speed.set("")
    read_packages(conn)


def create_customer(con):
    """This function is used to add a new customer to the table."""

    cursor = con.cursor()
    cursor.execute('SET IDENTITY_INSERT dbo.customers ON;'
                   'insert into dbo.customers(Customer_Id, First_Name, Last_Name, Birth_Date, Join_Date, City,'
                   'State, Street, main_phone_num, secondary_phone_num, fax, monthly_discount, pack_id)'
                   ' values(?,?,?,?,?,?,?,?,?,?,?,?,?);'
                   'SET IDENTITY_INSERT dbo.customers OFF;', (get_last_customer_id(con)+1,
                                                              client_stringVars["first_name"].get(),
                                                              client_stringVars["last_name"].get(),
                                                              client_stringVars["birth_date"].get(),
                                                              datetime.date.today().isoformat(),
                                                              client_stringVars["city"].get(),
                                                              client_stringVars["state"].get(),
                                                              client_stringVars["street"].get(),
                                                              client_stringVars["main_phone"].get(),
                                                              client_stringVars["sec_phone"].get(),
                                                              client_stringVars["fax"].get(),
                                                              client_stringVars["discount"].get(),
                                                              randomize_package_id(con)))
    con.commit()


def randomize_package_id(con):
    """This function is used to randomize a package id when assigning a package to a customer."""
    cursor = con.cursor()
    cursor.execute('SELECT * FROM Communications_LTD.dbo.packages')
    return random.randint(1, len(cursor.fetchall())-1)


def get_last_customer_id(con):
    """This function is used to retrieve the last customer id from the table in order to increment it by one when adding
     a new customer to the table."""
    cursor = con.cursor()
    cursor.execute('SELECT * FROM Communications_LTD.dbo.customers')
    return cursor.fetchall()[len(cursor.fetchall())-1][0]


def get_last_package_id(con):
    """This function is used to retrieve the last package id from the table in order to increment it by one when adding
    a new package to the table."""
    cursor = con.cursor()
    cursor.execute('SELECT * FROM Communications_LTD.dbo.packages')
    return cursor.fetchall()[len(cursor.fetchall())-1][0]


def create_package(con):
    """This function is used to add a new package to the table."""
    cursor = con.cursor()
    cursor.execute('SET IDENTITY_INSERT dbo.packages ON;'
                   'insert into dbo.packages(pack_id, speed, strt_date, monthly_payment, sector_id)'
                   ' values(?,?,?,?,?);'
                   'SET IDENTITY_INSERT dbo.packages OFF;', (get_last_package_id(con)+1,
                                                             package_stringVars["speed"].get(),
                                                             package_stringVars["start_date"].get(),
                                                             package_stringVars["monthly_payment"].get(),
                                                             package_stringVars["sector_id"].get()))
    con.commit()


def update_package(con):
    """This function is used when editing an existing package."""
    cursor = con.cursor()
    cursor.execute('update dbo.packages set speed = ? where pack_id = ?', (speed.get(), package_id.get()))
    con.commit()
    read_packages(con)


def menu_show():
    """This function is used to display the main menu.
    This loop is used to clear the window when coming back to the main menu."""
    for child in company.children.values():
        child.grid_forget()
    menu = "\n  1.Add a new client\n       2.Add a new package\n\t     3. Update an Existing Package"
    Label(company, text="Welcome to Communications LTD!", bg="sky blue", font="Times 20").grid(row=0, column=0)
    Label(company, text="------------------------------", bg="sky blue", font="Times 20").grid(row=1, column=0)
    for i in range(2, 7, 1):
        Label(company, text="                               ", bg="sky blue", font="Times 10").grid(row=i, column=0)
    Label(company, text="Please choose an option from the menu below :", bg="sky blue", font="Times 10").grid(row=7,
                                                                                                              column=0)
    Label(company, text=menu, bg="sky blue", font="Times 10").grid(row=8, column=0)
    Entry(company, textvariable=user_option, font="Times 10", width=10).grid(row=9, column=0)
    Button(company, command=main_menu_nav, text="submit", bg="yellow", font="Times 10").grid(row=9, column=1)
    Label(company, text="                               ", bg="sky blue", font="Times 10").grid(row=6, column=0)
    Button(company, command=company_out, text="Exit", bg="yellow", font="Times 10").grid(row=10, column=0)


def main_menu_nav():
    """This method is used to navigate between the different options of the menu and to clear the window after each
    choice. """
    if user_option.get() == 1:
        for child in company.children.values():
            child.grid_forget()
        add_client()
        user_option.set("")
    elif user_option.get() == 2:
        for child in company.children.values():
            child.grid_forget()
        add_package()
        user_option.set("")
    elif user_option.get() == 3:
        for child in company.children.values():
            child.grid_forget()
        edit_package()
        user_option.set("")


def company_out():
    company.destroy()


def edit_package():
    """This function is used to create a new window after the user selects an option from the main menu."""

    Label(company, text="Enter the package's id :", bg="sky blue", font="Times 10").grid(row=0, column=0)
    Label(company, text="                       ", bg="sky blue", font="Times 10").grid(row=1, column=0)
    Entry(company, textvariable=package_id, font="Times 10", width=10).grid(row=2, column=0)
    Label(company, text="                       ", bg="sky blue", font="Times 10").grid(row=3, column=0)
    Label(company, text="Enter the desired package speed id :", bg="sky blue", font="Times 10").grid(row=4, column=0)
    Label(company, text="                       ", bg="sky blue", font="Times 10").grid(row=5, column=0)
    Entry(company, textvariable=speed, font="Times 10", width=10).grid(row=6, column=0)
    Button(company, command=exeupdate_package, text="SUBMIT", bg="yellow", font="Times 10").grid(row=14, column=3)
    Button(company, command=menu_show, text="Exit to main menu", bg="yellow", font="Times 10").grid(row=15, column=0)


def add_client():
    """This function is used to display the new customer's entry fields after the user's selects option 1 from
     the main menu."""
    strings = ["Enter first name: ", "Enter last name : ", "Enter birth date : ",
               "Enter city: ", "Enter state: ", "Enter street: ", "Enter main phone : ", "Enter secondary phone : ",
               "Enter fax number : ", "Enter monthly discount : "]
    spacer = "                               "
    Label(company, text="\nEnter the new user's details: ", bg="sky blue", font="Times 10").grid(row=0, column=0)
    Label(company, text="                               ", bg="sky blue", font="Times 10").grid(row=1, column=0)
    index = 0
    for text in strings:
        Label(company, text=text, bg="sky blue", font="Times 10").grid(row=index+2, column=0)
        Label(company, text=spacer, bg="sky blue", font="Times 10").grid(row=index+3, column=0)
        index += 1

    index = 0
    for key in client_stringVars:
        Entry(company, textvariable=client_stringVars[key], font="Times 10", width=10).grid(row=index+2, column=1)
        index += 1

    Button(company, command=execreate_customer, text="SUBMIT", bg="yellow", font="Times 10").grid(row=14, column=3)
    Button(company, command=menu_show, text="Exit to main menu", bg="yellow", font="Times 10").grid(row=15, column=0)


def add_package():
    """This function is used to display the new package's entry fields after the user's selects option 2 from
     the main menu."""
    strings = ["Enter speed: ", "Enter start date  : ", "Enter monthly payment : ", "Enter sector id (1 or 2) : "]
    spacer = "                               "
    Label(company, text="\nEnter the new package's details: ", bg="sky blue", font="Times 10").grid(row=0, column=0)
    Label(company, text="                               ", bg="sky blue", font="Times 10").grid(row=1, column=0)
    index = 0

    for text in strings:
        Label(company, text=text, bg="sky blue", font="Times 10").grid(row=index + 2, column=0)
        Label(company, text=spacer, bg="sky blue", font="Times 10").grid(row=index + 3, column=0)
        index += 1
    index = 0
    for key in package_stringVars:
        Entry(company, textvariable=package_stringVars[key], font="Times 10", width=10).grid(row=index+2, column=1)
        index += 1

    Button(company, command=execreate_package, text="SUBMIT", bg="yellow", font="Times 10").grid(row=14, column=3)
    Button(company, command=menu_show, text="Exit to main menu", bg="yellow", font="Times 10").grid(row=15, column=0)


def read_customers(con):
    """This function is used to display all the existing customers from the customers table. """
    print("read")
    cursor = con.cursor()
    cursor.execute('SELECT * FROM Communications_LTD.dbo.customers')
    for row in cursor:
        print(f'row = {row}')
    print()


def read_packages(con):
    """This function is used to display all the existing packages from the packages table. """
    print("read")
    cursor = con.cursor()
    cursor.execute('SELECT * FROM Communications_LTD.dbo.packages')
    for row in cursor:
        print(f'row = {row}')
    print()

def stored_proc(conn):

    sqlCreateSP = "CREATE PROCEDURE pySelect_Records AS " \
                  "SELECT First_Name, Birth_Date " \
                  "FROM customers ORDER BY First_Name"

    sqlDropSP = "IF (OBJECT_ID('pySelect_Records', N'P') IS NOT NULL) BEGIN " \
                "DROP PROCEDURE pySelect_Records " \
                "END"

    sqlExecSP = "{call pySelect_Records }"

    cursor = conn.cursor()

    print("\nStored Procedure is : pySelect_Records")

    cursor.execute(sqlDropSP)

    cursor.execute(sqlCreateSP)

    try:
        cursor.execute(sqlExecSP)
    except pyodbc.Error as err:
        print('Error!!! %s' % err)
    print("\nResults :")

    recs = cursor.fetchall()

    for rec in recs:
        print("\nName : ", rec[0])
        print("\nAge : ", rec[1])

    cursor.close()

    del cursor


menu_show()
stored_proc(conn)
# company.mainloop()
# read_customers(conn)
# read_packages(conn)
