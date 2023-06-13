import tkinter as tk
import oracledb
from tkinter import messagebox, filedialog, ttk
import xml.etree.ElementTree as ET
import xml.dom.minidom
import datetime
import os


# POŁĄCZENIE ORACLE SQL
def connect_to_database(login_window, password):
    user = "s101450"
    host = "217.173.198.135"
    port = 1521
    service_name = "tpdb"

    try:
        dsn = oracledb.makedsn(host, port, service_name=service_name)
        connection = oracledb.connect(user=user, password=password, dsn=dsn)
        login_window.destroy()
        create_main_window(connection)
    except oracledb.DatabaseError as e:
        messagebox.showerror("Błąd", "Wystąpił błąd podczas nawiązywania połączenia: " + str(e))


# EXPORT TABEL DO XML
def export_table(table_name, root, connection):
    cursor = connection.cursor()
    try:
        cursor.execute("SELECT * FROM " + table_name)
        rows = cursor.fetchall()

        table_element = ET.SubElement(root, table_name)

        for row in rows:
            row_element = ET.SubElement(table_element, "row")

            for i, value in enumerate(row):
                column_name = cursor.description[i][0]
                column_type = cursor.description[i][1]
                column_element = ET.SubElement(row_element, column_name)

                if isinstance(value, datetime.datetime) or column_type == oracledb.DATETIME:
                    column_element.text = value.date().isoformat() if isinstance(value, datetime.datetime) else str(
                        value)
                else:
                    if isinstance(value, str) and value.startswith("TO_DATE(") and value.endswith(")"):
                        value = value[8:-1]
                    column_element.text = str(value)

    except oracledb.DatabaseError as e:
        messagebox.showerror("Błąd", "Wystąpił błąd podczas eksportowania tabeli " + table_name + ": " + str(e))
    finally:
        cursor.close()


# IMPORT TABEL Z XML
def import_tables(file_path, connection):
    imported_tables = []

    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        selected_tables = [table_listbox.get(index) for index in table_listbox.curselection()]
        not_imported_tables = []

        for table_name in selected_tables:
            if table_name in [table.tag for table in root]:
                imported_tables.append(table_name)
                table_element = root.find(table_name)
                import_table(table_name, table_element, connection)
            else:
                not_imported_tables.append(table_name)

        if not_imported_tables:
            not_imported_table_names = ", ".join(not_imported_tables)
            imported_table_names = ", ".join(imported_tables)
            messagebox.showwarning("Informacja", f"Następujące zaznaczone tabele nie zostały znalezione w pliku XML: {not_imported_table_names}\n"
                                                f"Zaimportowane tabele: {imported_table_names}")
        else:
            imported_table_names = ", ".join(imported_tables)
            messagebox.showinfo("Sukces", f"Tabele zostały zaimportowane poprawnie: {imported_table_names}")

    except ET.ParseError as e:
        messagebox.showerror("Błąd", "Wystąpił błąd podczas parsowania pliku XML: " + str(e))
    except oracledb.DatabaseError as e:
        messagebox.showerror("Błąd", "Wystąpił błąd podczas importowania tabel: " + str(e))

    return imported_tables


# IMPORT JEDNEJ TABELI Z XML
def import_table(table_name, table_element, connection):
    cursor = connection.cursor()
    try:
        disable_constraints(connection, table_name)
        cursor.execute("DELETE FROM " + table_name)

        for row_element in table_element:
            column_values = {}
            for column_element in row_element:
                column_name = column_element.tag
                column_value = column_element.text
                column_values[column_name] = column_value

            columns = ",".join(column_values.keys())
            values = ",".join([f"{parse_zero_value(value)}" for value in column_values.values()])
            insert_query = f"INSERT INTO {table_name} ({columns}) VALUES ({values})"
            insert_query = insert_query.replace("''", "'")
            print("INSERT QUERY:", insert_query)
            cursor.execute(insert_query)

        connection.commit()
        enable_constraints(connection, table_name)
    except oracledb.DatabaseError as e:
        connection.rollback()
        raise e
    finally:
        cursor.close()


# PARSOWANIE ZER I NONE
def parse_zero_value(value):
    if value == "0.0" or value == "None":
        return "NULL"
    elif value:
        return f"{parse_date(value)}"
    else:
        return "NULL"


# PARSOWANIE DATY
def parse_date(date_str):
    try:
        date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        return "TO_DATE('" + date.isoformat() + "', 'YYYY-MM-DD')"
    except ValueError:
        return "'" + date_str + "'"


# PRZYCISK IMPORT
def import_button_click(connection):
    selected_indices = table_listbox.curselection()

    if selected_indices:
        file_path = filedialog.askopenfilename(initialdir=get_pwd(), filetypes=[("XML Files", "*.xml")])
        if file_path:
            imported_tables = import_tables(file_path, connection)

            if imported_tables:
                table_names = ", ".join(imported_tables)
                # messagebox.showinfo("Sukces", "Tabele " + table_names + " zostały zaimportowane z pliku: " + file_path)
            else:
                messagebox.showinfo("Błąd", "Wystąpił problem podczas importowania tabel z pliku: " + file_path)


# PRZYCISK EXPORT
def export_button_click(connection, main_window):
    selected_indices = table_listbox.curselection()

    if selected_indices:
        default_file_name = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".xml"
        file_path = filedialog.asksaveasfilename(defaultextension=".xml", initialfile=default_file_name, filetypes=[("XML Files", "*.xml")])
        if file_path:
            exported_tables = []
            root = ET.Element("tables")

            table_order = [
                "DZIALY_PRODUKTOWE",
                "STANOWISKA",
                "PRACOWNICY",
                "FIRMY_ZAMAWIAJACE",
                "KLIENCI_DETALICZNI",
                "PRODUKTY",
                "ZAMOWIENIA",
                "ZAMOWIENIA_PRODUKTY",
                "FAKTURY",
                "DOSTAWY"
            ]

            total_tables = len(selected_indices)
            progress_bar = ttk.Progressbar(main_window, mode="determinate", maximum=total_tables)
            progress_bar.pack(pady=10)

            for i, table_name in enumerate(table_order):
                if table_name in [table_listbox.get(index) for index in selected_indices]:
                    exported_tables.append(table_name)
                    export_table(table_name, root, connection)
                    progress_bar["value"] = i + 1
                    main_window.update()

            progress_bar.destroy()

            xml_string = ET.tostring(root, encoding="utf-8")
            pretty_xml_string = xml.dom.minidom.parseString(xml_string).toprettyxml(indent="   ")

            with open(file_path, "w", encoding="utf-8") as xml_file:
                xml_file.write(pretty_xml_string)

            table_names = ", ".join(exported_tables)
            messagebox.showinfo("Sukces", "Tabele " + table_names + " zostały wyeksportowane do pliku: " + file_path)


# WYŁĄCZNIK INTEGRALNOŚCI
def disable_constraints(connection, table_name):
    cursor = connection.cursor()
    try:
        # nazwy ograniczeń
        cursor.execute("SELECT constraint_name FROM user_constraints WHERE table_name = :table_name", {"table_name": table_name})
        constraint_names = [row[0] for row in cursor.fetchall()]

        # wyłączenie tymczasowe ograniczeń
        for constraint_name in constraint_names:
            cursor.execute("SELECT table_name, constraint_name FROM user_constraints WHERE r_constraint_name = :constraint_name", {"constraint_name": constraint_name})
            dependent_constraints = cursor.fetchall()

            for dependent_table, dependent_constraint in dependent_constraints:
                cursor.execute("ALTER TABLE " + dependent_table + " DROP CONSTRAINT " + dependent_constraint)

            # wyłącz indeksy
            cursor.execute("SELECT index_name FROM user_constraints WHERE table_name = :table_name AND constraint_name = :constraint_name", {"table_name": table_name, "constraint_name": constraint_name})
            index_names = [row[0] for row in cursor.fetchall()]

            for index_name in index_names:
                if index_name is not None:
                    cursor.execute("ALTER INDEX " + index_name + " UNUSABLE")

            cursor.execute("ALTER TABLE " + table_name + " DISABLE CONSTRAINT " + constraint_name + " CASCADE")

        connection.commit()
    except oracledb.DatabaseError as e:
        connection.rollback()
        raise e
    finally:
        cursor.close()


# WŁĄCZNIK INTEGRALNOŚCI
def enable_constraints(connection, table_name):
    cursor = connection.cursor()
    try:
        # nazwy ograniczeń
        cursor.execute("SELECT constraint_name FROM user_constraints WHERE table_name = :table_name",
                       {"table_name": table_name})
        constraint_names = [row[0] for row in cursor.fetchall()]

        # włącz ograniczenia
        for constraint_name in constraint_names:
            cursor.execute("ALTER TABLE " + table_name + " ENABLE CONSTRAINT " + constraint_name)

        connection.commit()
    except oracledb.DatabaseError as e:
        connection.rollback()
        raise e
    finally:
        cursor.close()


# OBSŁUGA LISTY
def listbox_click(event):
    selected_indices = table_listbox.curselection()

    if selected_indices:
        table_listbox.selection_clear(0, tk.END)
        for index in selected_indices:
            table_listbox.selection_set(index)

    update_button_state()


# BLOKADA PRZYCISKÓW
def update_button_state():
    selected_indices = table_listbox.curselection()

    if selected_indices:
        export_button.config(state=tk.NORMAL)
        import_button.config(state=tk.NORMAL)
    else:
        import_button.config(state=tk.DISABLED)
        export_button.config(state=tk.DISABLED)


# OKNO LOGOWANIA
def create_login_window():
    login_window = tk.Tk()
    login_window.title("Logowanie")
    login_window.geometry("250x200")

    style = ttk.Style()
    style.configure("TButton", font=("Arial", 12))

    main_text = tk.Label(login_window, text=f"ORACLE SQL - XML IMP/EXP (v1.0)\n"
                         f"Autor: Paweł Siemiginowski\n"
                         f"s101450")
    main_text.pack(pady=10)
    password_label = tk.Label(login_window, text="Wprowadź hasło:")
    password_label.pack()

    password_entry = tk.Entry(login_window, show="*")
    password_entry.pack(pady=10)

    login_button = ttk.Button(login_window, text="Zaloguj",
                             command=lambda: connect_to_database(login_window, password_entry.get()))
    login_button.pack()

    login_window.mainloop()


# OKNO GŁÓWNE APKI
def create_main_window(connection):
    main_window = tk.Tk()
    main_window.title("Aplikacja do XML - ORACLE")
    main_window.geometry("800x600")

    table_frame = tk.Frame(main_window, bd=2, relief=tk.RAISED)
    table_frame.pack(pady=10)

    style = ttk.Style()
    style.configure("TButton", font=("Arial", 12))

    global table_listbox, import_button, export_button

    table_listbox = tk.Listbox(table_frame, selectmode=tk.MULTIPLE, exportselection=False, width=50, height=15, bd=0,
                               relief=tk.FLAT, font=("Arial", 10))
    table_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
    table_listbox.bind("<<ListboxSelect>>", listbox_click)

    table_scrollbar = tk.Scrollbar(table_frame, orient=tk.VERTICAL)
    table_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    table_listbox.config(yscrollcommand=table_scrollbar.set)
    table_scrollbar.config(command=table_listbox.yview)

    export_button = ttk.Button(main_window, text="Eksportuj XML", state=tk.DISABLED,
                               command=lambda: export_button_click(connection, main_window))
    export_button.pack()

    import_button = ttk.Button(main_window, text="Importuj XML", state=tk.DISABLED,
                               command=lambda: import_button_click(connection))
    import_button.pack()

    update_button_state()

    try:
        cursor = connection.cursor()
        cursor.execute("SELECT table_name FROM user_tables")
        tables = cursor.fetchall()
        for table in tables:
            table_listbox.insert(tk.END, table[0])

        cursor.close()
    except oracledb.DatabaseError as e:
        messagebox.showerror("Błąd", "Wystąpił błąd podczas pobierania listy tabel: " + str(e))

    main_window.mainloop()


# FOLDER
def get_pwd():
    default_folder = os.getcwd()
    return default_folder


# MAIN
if __name__ == "__main__":
    create_login_window()
