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

# ###########################################################################################################
#                                           ## M O D E L O ##
# ###########################################################################################################
# Base de Datos
def conexion():
    con = sqlite3.connect("isp.db")
    return con


def crear_tabla():
    con = conexion()
    cursor = con.cursor()
    sql = "CREATE TABLE IF NOT EXISTS datos(id INTEGER PRIMARY KEY AUTOINCREMENT, cliente TEXT, email TEXT, telefono TEXT, plan TEXT, activo TEXT, fecha TEXT)"
    cursor.execute(sql)
    con.commit()


try:
    conexion()
    crear_tabla()
except:
    print("Error.")


def abrir_archivo():
    tipo_archivo = (
        ("Excel files (*.xls)", "*.xls"),
        ("Excel files (*.xlsx)", "*.xlsx"),
        ("All files", "*.*"),
    )

    filename = fd.askopenfilename(
        title="Abrir archivo", initialdir="/", filetypes=tipo_archivo
    )


def guardar_archivo():
    tipo_archivo = (
        ("Excel files (*.xls)", "*.xls"),
        ("Excel files (*.xlsx)", "*.xlsx"),
        ("All files", "*.*"),
    )

    nom_archivo_g = fd.asksaveasfilename(
        title="Guardar archivo", initialdir="/", filetypes=tipo_archivo
    )


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
        print(row)
        mi_tv.insert(
            "",
            0,
            text=row[0],
            values=(row[1], row[2], row[3], row[4], row[5], row[6]),
        )


def alta(texto, id_modif, cliente, email, telefono, plan, activo, fecha, tree):
    global cursor_tree
    cadena = telefono
    patron = "[0-9]*$"
    if (
        re.match(patron, cadena)
        and cursor_tree["Cliente: "] != ""
        and cursor_tree["Email: "] != ""
        and cursor_tree["Teléfono: "] != 0
        and cursor_tree["Plan: "] != ""
        and cursor_tree["Activo: "] != ""
    ):
        con = conexion()
        cursor = con.cursor()
        if texto == "Agregar":
            data = (cliente, email, telefono, plan, activo, fecha)
            sql = "INSERT INTO datos(cliente, email, telefono, plan, activo, fecha) VALUES(?, ?, ?, ?, ?, ?)"
            cursor.execute(sql, data)
            con.commit()
            actualizar_treeview(tree)
        elif texto == "Modificar":
            data = (cliente, email, telefono, plan, activo, fecha, id_modif)

            sql = "UPDATE datos SET cliente=?, email=?, telefono=?, plan=?, activo=?, fecha=? WHERE id = ?"
            cursor.execute(sql, data)
            con.commit()
            actualizar_treeview(tree)
    else:
        print(
            "Error en campos.",
            "Se detectaron campos vacíos o erroneos, rellénelos correctamente para ingresar el cliente.",
        )


def f_baja():
    global cursor_tree, tree
    item = tree.item(tree.focus())
    id_borrar = list(item.values())[0]
    con = conexion()
    cursor = con.cursor()
    data = (id_borrar,)
    sql = "DELETE FROM datos WHERE id = ?"
    cursor.execute(sql, data)
    con.commit()
    actualizar_treeview(tree)


def f_consulta(tree):
    try:
        conexion()
        actualizar_treeview(tree)
    except:
        print("Error.")


def f_exportar():
    pass


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
var_cliente = StringVar()
var_email = StringVar()
var_telefono = StringVar()
var_plan = StringVar(value="Internet")
var_activo = StringVar(value="Si")
fecha = datetime.date.today()
fecha.strftime("%Y-%M-%D")

# Tuplas
t_plan = ("Internet", "IPTV", "Telefonía", "Hosting")
t_activo = ("Si", "No")

# Label
l_cliente = Label(root, text="Ingresar cliente:")
l_email = Label(root, text="Ingresar email:")
l_tel = Label(root, text="Ingresar teléfono:")
l_plan = Label(root, text="Ingresar plan:")
l_activo = Label(root, text="Activo: ")

# Entry
cliente = Entry(root, textvariable=var_cliente, width=35)
email = Entry(root, textvariable=var_email, width=35)
tel = Entry(root, textvariable=var_telefono, width=35)
# var_telefono.trace_add("write", lambda *args: limitador(var_telefono))

# ComboBox
plan = ttk.Combobox(
    root, textvariable=var_plan, width=35, state="readonly", values=t_plan
)
activo = ttk.Combobox(
    root,
    textvariable=var_activo,
    width=35,
    state="readonly",
    values=t_activo,
)

# Botones
b_agregar_cliente = Button(
    root,
    text="Agregar cliente",
    command=lambda: alta(
        var_cliente.get(),
        var_email.get(),
        var_telefono.get(),
        var_plan.get(),
        var_activo.get(),
        fecha,
        tree,
    ),
    width=15,
)
b_agregar_cliente.configure(border=3)

b_consulta = Button(root, text="Consulta", command=lambda: f_consulta(tree), width=15)
b_consulta.configure(border=3)

b_modificacion = Button(
    root, text="Modificación", command=lambda: ("Modificar"), width=15
)
b_modificacion.configure(border=3)

b_baja = Button(root, text="Baja", command=f_baja, width=15)
b_baja.configure(border=3)

b_exportar = Button(root, text="Export", command=f_exportar, width=15)
b_exportar.configure(border=3)

# Grid
l_cliente.grid(row=1, column=1, padx=10, pady=10, sticky=E)
cliente.grid(row=1, column=2, padx=10, pady=10)
l_email.grid(row=2, column=1, padx=10, pady=10, sticky=E)
email.grid(row=2, column=2, padx=10, pady=10)
l_tel.grid(row=3, column=1, padx=10, pady=10, sticky=E)
tel.grid(row=3, column=2, padx=10, pady=10)
l_plan.grid(row=4, column=1, padx=10, pady=10, sticky=E)
plan.grid(row=4, column=2, sticky=W, padx=10, pady=10)
l_activo.grid(row=5, column=1, padx=10, pady=10, sticky=E)
activo.grid(row=5, column=2, sticky=W, padx=10, pady=10)

b_agregar_cliente.grid(row=6, column=2, sticky=W, padx=20, pady=10)
b_agregar_cliente.grid_propagate(True)
b_consulta.grid(row=8, column=1, sticky=W, padx=20, pady=10)
b_consulta.grid_propagate(True)
b_modificacion.grid(row=8, column=2, sticky=W, padx=20, pady=10)
b_modificacion.grid_propagate(True)
b_baja.grid(row=8, column=3, sticky=W, padx=20, pady=10)
b_baja.grid_propagate(True)
b_exportar.grid(row=8, column=4, sticky=W, padx=20, pady=10)
b_exportar.grid_propagate(True)

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
tree.column("Cliente", width=150, minwidth=150, anchor=W)
tree.column("Email", width=200, minwidth=200, anchor=W)
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


root.mainloop()
