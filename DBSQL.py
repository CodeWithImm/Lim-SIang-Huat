import csv
import logging
import pymysql
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
from ttkthemes import ThemedStyle

DATABASE_CONFIG_FILE = "config.ini"
LOG_FILE = "app_log.log"
DB_HOST = ""
DB_USERNAME = ""
DB_PASSWORD = ""
DB_DATABASE = ""
undo_stack = []
redo_stack = []

logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def read_database_config():
    global DB_HOST, DB_USERNAME, DB_PASSWORD, DB_DATABASE
    try:
        with open(DATABASE_CONFIG_FILE, "r") as config_file:
            config_data = config_file.readlines()
            for line in config_data:
                key, value = line.strip().split("=")
                if key == "host":
                    DB_HOST = value
                elif key == "username":
                    DB_USERNAME = value
                elif key == "password":
                    DB_PASSWORD = value
                elif key == "database":
                    DB_DATABASE = value
    except FileNotFoundError:
        messagebox.showwarning("Warning", "Database configuration file not found.")
        logging.warning("Database configuration file not found.")

def connect_to_database():
    try:
        mydb = pymysql.connect(host=DB_HOST, user=DB_USERNAME, password=DB_PASSWORD, database=DB_DATABASE)
        return mydb
    except pymysql.Error as e:
        logging.error(f"Error connecting to database: {e}")
        messagebox.showerror("Error", f"Error connecting to database: {e}")
        return None

def save_database_config():
    global DB_HOST, DB_USERNAME, DB_PASSWORD, DB_DATABASE
    with open(DATABASE_CONFIG_FILE, "w") as config_file:
        config_file.write(f"host={DB_HOST}\n")
        config_file.write(f"username={DB_USERNAME}\n")
        config_file.write(f"password={DB_PASSWORD}\n")
        config_file.write(f"database={DB_DATABASE}\n")
    messagebox.showinfo("Settings Saved", "Database settings saved successfully.")
    logging.info("Database settings saved successfully.")

def test_database_connection():
    if connect_to_database():
        messagebox.showinfo("Connection Test", "Database connection successful!")
        logging.info("Database connection successful!")
    else:
        messagebox.showerror("Connection Test", "Failed to connect to the database.")
        logging.error("Failed to connect to the database.")

def open_settings():
    def save_settings():
        global DB_HOST, DB_USERNAME, DB_PASSWORD, DB_DATABASE
        DB_HOST = host_entry.get()
        DB_USERNAME = username_entry.get()
        DB_PASSWORD = password_entry.get()
        DB_DATABASE = database_entry.get()
        save_database_config()

    def test_connection():
        test_database_connection()

    settings_window = tk.Toplevel(root)
    settings_window.title("Database Settings")

    ttk.Label(settings_window, text="Host:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    ttk.Label(settings_window, text="Username:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    ttk.Label(settings_window, text="Password:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
    ttk.Label(settings_window, text="Database:").grid(row=3, column=0, padx=10, pady=5, sticky="e")

    host_entry = ttk.Entry(settings_window)
    host_entry.grid(row=0, column=1, padx=10, pady=5)
    host_entry.insert(0, DB_HOST)
    username_entry = ttk.Entry(settings_window)
    username_entry.grid(row=1, column=1, padx=10, pady=5)
    username_entry.insert(0, DB_USERNAME)
    password_entry = ttk.Entry(settings_window, show="*")
    password_entry.grid(row=2, column=1, padx=10, pady=5)
    password_entry.insert(0, DB_PASSWORD)
    database_entry = ttk.Entry(settings_window)
    database_entry.grid(row=3, column=1, padx=10, pady=5)
    database_entry.insert(0, DB_DATABASE)

    ttk.Button(settings_window, text="Save", command=save_settings).grid(row=4, column=1, pady=10)
    ttk.Button(settings_window, text="Test Connection", command=test_connection).grid(row=4, column=0, pady=10)

def preview_files(data_type):
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if filename:
        try:
            with open(filename, newline='') as csvfile:
                reader = csv.reader(csvfile, delimiter=';')
                data = list(reader)

                headers = None
                if data_type == "SBC": 
                    headers = data[0]
                elif data_type == "SBO":  
                    headers = data[0][:6] + data[0][7:12]  
                elif data_type == "SISO":  
                    headers = data[0][:13] 

                if headers:
                    tree.delete(*tree.get_children())

                    
                    for idx, row in enumerate(data):
                        if idx == 0: 
                            tree["columns"] = headers
                            tree.heading("#0", text="Index")
                            for i, col in enumerate(headers):
                                tree.column(col, width=100)
                                tree.heading(col, text=col)
                        else:
                            tree.insert("", "end", text=idx, values=row)

                    total_records = len(data) - 1  
                    total_quantity = sum(float(row[9].replace(",", "")) for row in data[1:] if len(row) > 9 and row[9].replace(",", "").replace(".", "").isdigit())  # Summing up Qtytarget if row has enough elements

                    total_records_label.config(text=f"Total Records: {total_records}")
                    total_quantity_label.config(text=f"Total Quantity: {total_quantity}")

                
                    if total_records == 0:
                        save_button.config(state="disabled")
                    else:
                        save_button.config(state="normal")

                    
                    create_missing_table(data_type, headers)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading CSV file: {str(e)}")
            logging.error(f"Error loading CSV file: {str(e)}")

def create_missing_table(data_type, headers):
    table_name = f"tbl_{data_type.lower()}"
    column_definitions = ", ".join(f"{header} VARCHAR(255)" for header in headers)
    try:
        mydb = connect_to_database()
        if mydb:
            cursor = mydb.cursor()
            cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} ({column_definitions})")
            mydb.commit()
            logging.info(f"Table {table_name} checked/created successfully.")
            cursor.close()
            mydb.close()
    except pymysql.Error as e:
        messagebox.showerror("Error", f"An error occurred while checking/creating table: {str(e)}")
        logging.error(f"Error checking/creating table: {str(e)}")

def undo():
    global undo_stack, redo_stack
    if undo_stack:
        redo_stack.append(undo_stack.pop())  
        tree.delete(*tree.get_children())  
        for item in undo_stack[-1]:  
            tree.insert("", "end", values=item)

def redo():
    global undo_stack, redo_stack
    if redo_stack:
        undo_stack.append(redo_stack.pop())
        tree.delete(*tree.get_children())  
        for item in undo_stack[-1]:  
            tree.insert("", "end", values=item)

def save_filtered_data_to_mysql():
    filtered_data = []
    for item in tree.get_children():
        values = tree.item(item, "values")
        filtered_data.append(values)

    save_to_mysql(filtered_data)

def insert_data():
    try:
        mydb = connect_to_database()
        if mydb:
            cursor = mydb.cursor()

            selected_items = tree.selection()
            if not selected_items:  
                messagebox.showerror("Error", "Please select an item to insert.")
                return
           
            for selected_item in selected_items:
                values = tree.item(selected_item, "values")
                csv_type = values[0].split(';')[0].lower()
                
                if csv_type == 'sbc':
                    query = "INSERT INTO tbl_sbc (Region_Name, Customer_Group, Customer_Name, Year_of_Date, Month_of_Date, Day_of_Date, Channel_Class_Code, Product_Brand, Product_Name, Qtytarget, So_Qty, So_ValueTarget, So_Value, date1) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                elif csv_type == 'sbo':
                    query = "INSERT INTO tbl_sbo (Region_Name, Customer_Group, Year_of_Date, Month_of_Date, Customer_Name, Product_Brand, Product_Name, Channel_Class_Code, Outlet_Code, Outlet_Name, Sales_Code, Channel_Category_All, QTY, VALUE, date1, date2, Location, Day_of_Date) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                elif csv_type == 'siso':
                    query = "INSERT INTO tbl_siso (Region_Name, Year_of_Date, Month_of_Date, Customer_Name, Product_Brand, Product_Name, SO_VALUETARGET, Valuetarget, So_Value, Si_Value, Si_Qty, So_Qty, Qtytarget, date1, date2, Day_of_Date) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                
                if len(values[1:]) == query.count('%s'):
                    cursor.execute(query, values[1:]) 
                else:
                    messagebox.showerror("Error", "Number of values does not match the expected number of columns.")
            
            mydb.commit()
            messagebox.showinfo("Success", "Data inserted successfully!")
            cursor.close()
            mydb.close()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while inserting data: {str(e)}")
        logging.error(f"Error inserting data: {str(e)}")

def update_data():
    try:
        mydb = connect_to_database()
        if mydb:
            cursor = mydb.cursor()
            selected_item = tree.selection()
            if not selected_item:  
                messagebox.showerror("Error", "Please select an item to update.")
                return
            selected_item = selected_item[0]
            values = tree.item(selected_item, "values")
            
            csv_type = values[0].split()[0].lower()
            if csv_type == 'sbc':
                query = "UPDATE tbl_sbc SET Region_Name = %s, Customer_Group = %s, Customer_Name = %s, Year_of_Date = %s, Month_of_Date = %s, Day_of_Date = %s, Channel_Class_Code = %s, Product_Brand = %s, Product_Name = %s, Qtytarget = %s, So_Qty = %s, So_ValueTarget = %s, So_Value = %s, date1 = %s WHERE id = %s"
            elif csv_type == 'sbo':
                query = "UPDATE tbl_sbo SET Region_Name = %s, Customer_Group = %s, Year_of_Date = %s, Month_of_Date = %s, Customer_Name = %s, Product_Brand = %s, Product_Name = %s, Channel_Class_Code = %s, Outlet_Code = %s, Outlet_Name = %s, Sales_Code = %s, Channel_Category_All = %s, QTY = %s, VALUE = %s, date1 = %s, date2 = %s, Location = %s, Day_of_Date = %s WHERE id = %s"
            elif csv_type == 'siso':
                query = "UPDATE tbl_siso SET Region_Name = %s, Year_of_Date = %s, Month_of_Date = %s, Customer_Name = %s, Product_Brand = %s, Product_Name = %s, SO_VALUETARGET = %s, Valuetarget = %s, So_Value = %s, Si_Value = %s, Si_Qty = %s, So_Qty = %s, Qtytarget = %s, date1 = %s, date2 = %s, Day_of_Date = %s WHERE id = %s"
            
            if len(values[1:]) == query.count('%s'):
                cursor.execute(query, values[1:] + [values[0]])  
            else:
                messagebox.showerror("Error", "Number of values does not match the expected number of columns.")
            
            mydb.commit()  
            messagebox.showinfo("Success", "Data updated successfully!")
            cursor.close()
            mydb.close()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while updating data: {str(e)}")
        logging.error(f"Error updating data: {str(e)}")


root = tk.Tk()
root.title("CSV to MySQL")


menubar = Menu(root)
file_menu = Menu(menubar, tearoff=0)
file_menu.add_command(label="Open Settings", command=open_settings)
file_menu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=file_menu)
root.config(menu=menubar)

main_frame = ttk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

tree = ttk.Treeview(main_frame)
tree.pack(side="left", fill="both", expand=True)

vsb = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
vsb.pack(side="right", fill="y")
tree.configure(yscrollcommand=vsb.set)

hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=tree.xview)
hsb.pack(side="bottom", fill="x")
tree.configure(xscrollcommand=hsb.set)

button_frame = ttk.Frame(root)
button_frame.pack(pady=10)

open_button = ttk.Button(button_frame, text="Open SBC File", command=lambda: preview_files("SBC"))
open_button.grid(row=0, column=0, padx=5)

open_button = ttk.Button(button_frame, text="Open SBO File", command=lambda: preview_files("SBO"))
open_button.grid(row=0, column=1, padx=5)

open_button = ttk.Button(button_frame, text="Open SISO File", command=lambda: preview_files("SISO"))
open_button.grid(row=0, column=2, padx=5)

undo_button = ttk.Button(button_frame, text="Undo", command=undo)
undo_button.grid(row=0, column=3, padx=5)

redo_button = ttk.Button(button_frame, text="Redo", command=redo)
redo_button.grid(row=0, column=4, padx=5)

save_button = ttk.Button(button_frame, text="Save to MySQL", command=save_filtered_data_to_mysql, state="disabled")
save_button.grid(row=0, column=5, padx=5)

insert_button = ttk.Button(button_frame, text="Insert to MySQL", command=insert_data)
insert_button.grid(row=0, column=6, padx=5)

update_button = ttk.Button(button_frame, text="Update to MySQL", command=update_data)
update_button.grid(row=0, column=7, padx=5)

total_records_label = ttk.Label(button_frame, text="Total Records: 0")
total_records_label.grid(row=1, column=0, columnspan=4)

total_quantity_label = ttk.Label(button_frame, text="Total Quantity: 0")
total_quantity_label.grid(row=1, column=4, columnspan=4)

read_database_config()

root.mainloop()
