from msilib.schema import ComboBox
from tkinter import *
from tkinter import ttk
from PIL import ImageTk, Image
from tkinter import filedialog as fd
from tkinter.messagebox import *
import os
import datetime
import sqlite3
import re
import csv

# 1° pip install xlsxwriter para poder manejar archivos de excel
import xlsxwriter

# ###########################################################################################################
#                                           ## M O D E L O ##
# ###########################################################################################################
# Base de Datos
def conexion():
    ruta = os.path.dirname(os.path.abspath(__file__))
    db = os.path.join(ruta, "isp.db")
    con = sqlite3.connect(db)
    return con


def crear_tabla():
    con = conexion()
    cursor = con.cursor()
    sql = "CREATE TABLE IF NOT EXISTS datos(id INTEGER PRIMARY KEY AUTOINCREMENT, cliente TEXT, email TEXT, telefono TEXT, plan TEXT, activo TEXT, fecha TEXT)"
    cursor.execute(sql)
    con.commit()


# Función que limita la cantidad de caracteres y que sólo permite ingresar números en el campo del teléfono
def limitador(entry_text):
    patron = "[0-9]*$"
    if re.match(patron, entry_text.get()):
        if len(entry_text.get()) > 0:
            # donde esta el :15 limitas la cantidad de caracteres
            entry_text.set(entry_text.get()[:15])
    else:
        showerror(
            "Error de caracter",
            "Tipo de caracter ingresado inválido, ingrese sólo números",
        )
        telefono.set(0)


# Función para limpiar los campos de los entry
def limpiar_campos():
    global cliente, email, telefono, plan, activo
    cliente.set("")
    email.set("")
    telefono.set(0)
    plan.set("")
    activo.set("")


def abrir_archivo():
    tipo_archivo = (
        ("Excel files (*.xls)", "*.xls"),
        ("Excel files (*.xlsx)", "*.xlsx"),
        ("All files", "*.*"),
    )

    filename = fd.askopenfilename(
        title="Abrir archivo", initialdir="/", filetypes=tipo_archivo
    )


# Función para guardar en un directorio particular un excel con los datos de la DB
def guardar_archivo():
    tipo_archivo = (
        ("Excel files (*.xls)", "*.xls"),
        ("Excel files (*.xlsx)", "*.xlsx"),
        ("All files", "*.*"),
    )

    ruta = os.path.dirname(os.path.abspath(__file__))
    nom_archivo_g = fd.asksaveasfilename(
        title="Guardar archivo", initialdir=ruta, filetypes=tipo_archivo
    )
    arch_xlsx = os.path.join(ruta, nom_archivo_g + ".xlsx")
    hoja = archivo.add_worksheet(name="nombre de hoja")
    sql = "SELECT * FROM datos ORDER BY id ASC"
    con = conexion()
    cursor = con.cursor()
    datos = cursor.execute(sql)
    result = datos.fetchall()

    hoja.write(0, 0, "N° de cliente")
    hoja.write(0, 1, "Cliente")
    hoja.write(0, 2, "Email")
    hoja.write(0, 3, "Teléfono")
    hoja.write(0, 4, "Plan")
    hoja.write(0, 5, "Activo")
    hoja.write(0, 6, "Fecha de creación")
    for row in range(len(result)):
        for col in range(len(result[row])):
            hoja.write(row + 1, col, result[row][col])

    archivo.close()


# Función que actualiza el treeview
def actualizar_treeview(mi_tv):
    records = mi_tv.get_children()
    for element in records:
        mi_tv.delete(element)
    sql = "SELECT * FROM datos ORDER BY id ASC"
    con = conexion()
    cursor = con.cursor()
    datos = cursor.execute(sql)
    result = datos.fetchall()
    for row in result:
        mi_tv.insert(
            "",
            0,
            text=row[0],
            values=(row[1], row[2], row[3], row[4], row[5], row[6]),
        )


# Función para dar de alta un cliente
def alta():
    global tree, cliente, email, telefono, plan, activo, fecha
    if (
        cliente.get() != ""
        and email.get() != ""
        and telefono.get() != 0
        and plan.get() != ""
        and activo.get() != ""
    ):
        con = conexion()
        cursor = con.cursor()
        data = (
            cliente.get(),
            email.get(),
            telefono.get(),
            plan.get(),
            activo.get(),
            fecha,
        )
        sql = "INSERT INTO datos(cliente, email, telefono, plan, activo, fecha) VALUES(?, ?, ?, ?, ?, ?)"
        cursor.execute(sql, data)
        con.commit()
        actualizar_treeview(tree)
        limpiar_campos()
        showinfo("Carga de datos", "Datos cargados correctamente.")
        """
            data = (cliente, email, telefono, plan, activo, fecha, id_modif)

            sql = "UPDATE datos SET cliente=?, email=?, telefono=?, plan=?, activo=?, fecha=? WHERE id = ?"
            cursor.execute(sql, data)
            con.commit()
            actualizar_treeview(tree)
            """
    else:
        showerror(
            "Error en campos",
            "Se detectaron campos vacíos o erroneos, rellénelos correctamente para ingresar el cliente.",
        )


# Función para dar de baja un cliente
def f_baja():
    global tree, cliente, email, telefono, plan, activo, fecha
    item = tree.item(tree.focus())
    id_borrar = list(item.values())[0]
    con = conexion()
    cursor = con.cursor()
    data = (id_borrar,)
    sql = "DELETE FROM datos WHERE id = ?"
    cursor.execute(sql, data)
    con.commit()
    actualizar_treeview(tree)
    limpiar_campos()
    showinfo("Baja de cliente", "Datos eliminados correctamente.")


# Función que carga los entry con el valor del item seleccionado
def itemseleccionado(event):
    global tree, mi_id, cliente, email, telefono, plan, activo, fecha
    item_sel = tree.item(tree.focus())
    lista_seleccionada = list(item_sel.values())
    mi_id.set(lista_seleccionada[0])
    cliente.set(lista_seleccionada[2][0])
    email.set(lista_seleccionada[2][1])
    telefono.set(lista_seleccionada[2][2])
    plan.set(lista_seleccionada[2][3])
    activo.set(lista_seleccionada[2][4])
    fecha = lista_seleccionada[2][5]


def f_modificar():
    global tree, cliente, email, telefono, plan, activo, fecha, mi_id
    if (
        cliente.get() != ""
        and email.get() != ""
        and telefono.get() != 0
        and plan.get() != ""
        and activo.get() != ""
    ):
        con = conexion()
        cursor = con.cursor()
        data = (
            cliente.get(),
            email.get(),
            telefono.get(),
            plan.get(),
            activo.get(),
            fecha,
            mi_id.get(),
        )
        sql = "UPDATE datos SET cliente=?, email=?, telefono=?, plan=?, activo=?, fecha=? WHERE id = ?"
        cursor.execute(sql, data)
        con.commit()
        actualizar_treeview(tree)
        showinfo("Carga de datos", "Datos cargados correctamente.")
    else:
        showerror(
            "Error en campos",
            "Se detectaron campos vacíos o erroneos, rellénelos correctamente para ingresar el cliente.",
        )


def f_consulta():
    global tree
    try:
        conexion()
        actualizar_treeview(tree)
    except:
        print("Error.")


def f_exportar():
    ruta = os.path.dirname(os.path.abspath(__file__))
    arch_xlsx = os.path.join(ruta, "export_db.xlsx")
    archivo = xlsxwriter.Workbook(arch_xlsx)
    hoja = archivo.add_worksheet(name="nombre de hoja")
    sql = "SELECT * FROM datos ORDER BY id ASC"
    con = conexion()
    cursor = con.cursor()
    datos = cursor.execute(sql)
    result = datos.fetchall()
    hoja.write(0, 0, "N° de cliente")
    hoja.write(0, 1, "Cliente")
    hoja.write(0, 2, "Email")
    hoja.write(0, 3, "Teléfono")
    hoja.write(0, 4, "Plan")
    hoja.write(0, 5, "Activo")
    hoja.write(0, 6, "Fecha de creación")
    for row in range(len(result)):
        for col in range(len(result[row])):
            hoja.write(row + 1, col, result[row][col])
    archivo.close()
    showinfo("Exito", "Archivo Excel guardado correctamente")


try:
    conexion()
    crear_tabla()
except:
    print("Error.")

# ###########################################################################################################
#                                           ## V I S T A ##
# ###########################################################################################################
# Root
root = Tk()
root.title("Proyecto ISP")
root.configure(background="#ADE7FF")
root.geometry("1000x600")
root.resizable(width=False, height=False)

# Icono
BASE_DIR = os.path.dirname((os.path.abspath(__file__)))
ruta = os.path.join(BASE_DIR, "img", "icono.png")
image_root = Image.open(ruta)
image_icon = ImageTk.PhotoImage(image_root)
root.iconphoto(True, image_icon)

# Foto barra
rute = os.path.join(BASE_DIR, "img", "logo.png")
img = Image.open(rute)
n_img = img.resize((1000, 50))
render = ImageTk.PhotoImage(n_img)
img1 = Label(root, image=render)
img1.image = render
img1.grid(row=0, column=0, columnspan=8)

# Variables
mi_id = StringVar()
cliente = StringVar()
email = StringVar()
telefono = IntVar()
plan = StringVar(value="Internet")
activo = StringVar(value="Si")
fecha = datetime.date.today()
fecha.strftime("%Y-%M-%D")

# Tuplas
t_plan = ("Internet", "IPTV", "Telefonía", "Hosting")
t_activo = ("Si", "No")

# Label
l_cliente = Label(root, text="Ingresar cliente:", background="#ADE7FF")
l_email = Label(root, text="Ingresar email:", background="#ADE7FF")
l_tel = Label(root, text="Ingresar teléfono:", background="#ADE7FF")
l_plan = Label(root, text="Ingresar plan:", background="#ADE7FF")
l_activo = Label(root, text="Activo: ", background="#ADE7FF")

# Entry
e_cliente = Entry(root, textvariable=cliente, width=35)
e_email = Entry(root, textvariable=email, width=35)
e_tel = Entry(root, textvariable=telefono, width=35)

# ComboBox
cmb_plan = ttk.Combobox(
    root, textvariable=plan, width=35, state="readonly", values=t_plan
)
cmb_activo = ttk.Combobox(
    root,
    textvariable=activo,
    width=35,
    state="readonly",
    values=t_activo,
)

# Botones
b_agregar_cliente = Button(root, text="Agregar cliente", command=alta, width=15)
b_agregar_cliente.configure(border=3)

b_consulta = Button(root, text="Consulta", command=f_consulta, width=15)
b_consulta.configure(border=3)

b_modificacion = Button(root, text="Modificación", command=f_modificar, width=15)
b_modificacion.configure(border=3)

b_baja = Button(root, text="Baja", command=f_baja, width=15)
b_baja.configure(border=3)

b_exportar = Button(root, text="Export", command=f_exportar, width=15)
b_exportar.configure(border=3)

# Grid
l_cliente.grid(row=1, column=1, padx=10, pady=10, sticky=E)
e_cliente.grid(row=1, column=2, padx=10, pady=10, sticky=W)
l_email.grid(row=2, column=1, padx=10, pady=10, sticky=E)
e_email.grid(row=2, column=2, padx=10, pady=10, sticky=W)
l_tel.grid(row=3, column=1, padx=10, pady=10, sticky=E)
e_tel.grid(row=3, column=2, padx=10, pady=10, sticky=W)
l_plan.grid(row=4, column=1, padx=10, pady=10, sticky=E)
cmb_plan.grid(row=4, column=2, sticky=W, padx=10, pady=10)
l_activo.grid(row=5, column=1, padx=10, pady=10, sticky=E)
cmb_activo.grid(row=5, column=2, sticky=W, padx=10, pady=10)
b_agregar_cliente.grid(row=6, column=2, sticky=W, padx=20, pady=10)
b_consulta.grid(row=8, column=1, sticky=W, padx=20, pady=10)
b_modificacion.grid(row=8, column=2, sticky=W, padx=20, pady=10)
b_baja.grid(row=8, column=3, sticky=W, padx=20, pady=10)
b_exportar.grid(row=8, column=4, sticky=W, padx=20, pady=10)

# MenuBar
menubar = Menu(root)

menu_archivo = Menu(menubar, tearoff=0)
menu_archivo.add_command(label="Abrir", command=abrir_archivo)
menu_archivo.add_command(label="Guardar como...", command=guardar_archivo)
menu_archivo.add_separator()
menu_archivo.add_command(label="Salir", command=root.quit)
menubar.add_cascade(label="Archivo", menu=menu_archivo)

root.config(menu=menubar)

# Treview
tree = ttk.Treeview(root)
tree["columns"] = (
    "Cliente",
    "Email",
    "Teléfono",
    "Plan",
    "Activo",
    "Fecha de creación",
)
# Configuración de las columnas
tree.column("#0", width=90, minwidth=90, anchor=W)
tree.column("Cliente", width=150, minwidth=150, anchor=CENTER)
tree.column("Email", width=200, minwidth=200, anchor=CENTER)
tree.column("Teléfono", width=100, minwidth=100, anchor=E)
tree.column("Plan", width=180, minwidth=180, anchor=E)
tree.column("Activo", width=50, minwidth=50, anchor=E)
tree.column("Fecha de creación", width=180, minwidth=180, anchor=E)
# Título de las columnas
tree.heading("#0", text="N° de cliente")
tree.heading("#1", text="Cliente")
tree.heading("#2", text="Email")
tree.heading("#3", text="Teléfono")
tree.heading("#4", text="Plan")
tree.heading("#5", text="Activo")
tree.heading("#6", text="Fecha de creación")
tree.grid(row=7, column=0, columnspan=7, padx=20, pady=10)
tree.bind(sequence="<<TreeviewSelect>>", func=itemseleccionado)

root.mainloop()
