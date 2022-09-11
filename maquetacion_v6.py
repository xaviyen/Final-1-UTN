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

mi_id = 0
cursor_tree = {}
#################
# --Funciones-- #
#################
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

# Función que abre diálogo para abrir archivos
def abrir_archivo():
    tipo_archivo = (
        ("Excel files (*.xls)", "*.xls"),
        ("Excel files (*.xlsx)", "*.xlsx"),
        ("All files", "*.*"),
    )

    filename = fd.askopenfilename(
        title="Abrir archivo", initialdir="/", filetypes=tipo_archivo
    )


# Función que abre diálogo para guardar archivos
def guardar_archivo():
    tipo_archivo = (
        ("Excel files (*.xls)", "*.xls"),
        ("Excel files (*.xlsx)", "*.xlsx"),
        ("All files", "*.*"),
    )

    nom_archivo_g = fd.asksaveasfilename(
        title="Guardar archivo", initialdir="/", filetypes=tipo_archivo
    )


# Funcion para guardar en base de datos
def alta(cliente, email, telefono, plan, activo, tree):
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
        data = (cliente, email, telefono, plan, activo)
        sql = "INSERT INTO datos(cliente, email, telefono, plan, activo) VALUES(?, ?, ?, ?, ?)"
        cursor.execute(sql, data)
        con.commit()
        actualizar_treeview(tree)
    else:
        print(
            "Error en campos.",
            "Se detectaron campos vacíos o erroneos, rellénelos correctamente para ingresar el cliente.",
        )


# Funcion para guardar en treeview lo mismo que se guarda en base de datos
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
            values=(row[1], row[2], row[3], row[4], row[5], datetime.date.today()),
        )


# Ventana para agregar cliente
def f_agregar_cliente():
    global cursor_tree, tree, mi_id

    # Función que inserta los datos ingresados en el treeview principal
    def ingresar_datos_tree():
        global cursor_tree, tree, mi_id
        nonlocal nueva_ventana
        cadena = var_telefono.get()
        patron = "[0-9]*$"
        cursor_tree = {
            "Cliente: ": var_cliente.get(),
            "Email: ": var_email.get(),
            "Teléfono: ": var_telefono.get(),
            "Plan: ": var_plan.get(),
            "Activo: ": var_activo.get(),
        }

        if (
            re.match(patron, cadena)
            and cursor_tree["Cliente: "] != ""
            and cursor_tree["Email: "] != ""
            and cursor_tree["Teléfono: "] != 0
            and cursor_tree["Plan: "] != ""
            and cursor_tree["Activo: "] != ""
        ):
            showinfo("Exito", "¡Sus datos se han guardado correctamente!")
            nueva_ventana.destroy()
        else:
            showerror(
                "Error en campos.",
                "Se detectaron campos vacíos o erroneos, rellénelos correctamente para ingresar el cliente.",
            )

    # Declaración de variables
    var_cliente = StringVar()
    var_email = StringVar()
    var_telefono = StringVar()
    var_plan = StringVar()
    var_activo = StringVar()
    global cursor_tree, mi_id

    # Creación de la nueva ventana
    nueva_ventana = Toplevel(root)
    nueva_ventana.title("Agregar Cliente")
    nueva_ventana.geometry("450x265")
    nueva_ventana.focus()
    nueva_ventana.grab_set()  # -->Para bloquear la ventana principal

    # Tuplas para combobox
    t_plan = ("Internet", "IPTV", "Telefonía", "Hosting")
    t_activo = ("Si", "No")
    ####################################
    # --LABEL--ENTRY--GRID--COMBOBOX-- #
    ####################################
    # LABEL
    l_cliente = Label(nueva_ventana, text="Ingresar cliente:")
    l_email = Label(nueva_ventana, text="Ingresar email:")
    l_tel = Label(nueva_ventana, text="Ingresar teléfono:")
    l_plan = Label(nueva_ventana, text="Ingresar plan:")
    l_activo = Label(nueva_ventana, text="Activo: ")

    # ENTRY
    cliente = Entry(nueva_ventana, textvariable=var_cliente, width=40)
    email = Entry(nueva_ventana, textvariable=var_email, width=40)
    tel = Entry(nueva_ventana, textvariable=var_telefono, width=40)

    # COMBOBOX
    plan = ttk.Combobox(
        nueva_ventana, textvariable=var_plan, width=35, state="readonly", values=t_plan
    )
    activo = ttk.Combobox(
        nueva_ventana,
        textvariable=var_activo,
        width=35,
        state="readonly",
        values=t_activo,
    )

    # BUTTON
    b_aceptar_ingreso = Button(
        nueva_ventana,
        text="Aceptar",
        command=lambda: [
            ingresar_datos_tree(),
            alta(
                var_cliente.get(),
                var_email.get(),
                var_telefono.get(),
                var_plan.get(),
                var_activo.get(),
                tree,
            ),
        ],
    )

    b_cancelar_ingreso = Button(
        nueva_ventana, text="Cancelar", command=nueva_ventana.destroy
    )

    # GRID
    l_cliente.grid(row=0, column=0, padx=10, pady=10, sticky=E)
    cliente.grid(row=0, column=1, padx=10, pady=10)
    l_email.grid(row=1, column=0, padx=10, pady=10, sticky=E)
    email.grid(row=1, column=1, padx=10, pady=10)
    l_tel.grid(row=2, column=0, padx=10, pady=10, sticky=E)
    tel.grid(row=2, column=1, padx=10, pady=10)
    l_plan.grid(row=3, column=0, padx=10, pady=10, sticky=E)
    plan.grid(row=3, column=1, sticky=W, padx=10, pady=10)
    l_activo.grid(row=4, column=0, padx=10, pady=10, sticky=E)
    activo.grid(row=4, column=1, sticky=W, padx=10, pady=10)
    b_aceptar_ingreso.grid(row=5, column=0)
    b_cancelar_ingreso.grid(row=5, column=1)


def f_baja():
    pass


def f_modificacion():
    pass


def f_consulta():
    pass


def f_exportar():
    pass


root = Tk()
# Ajuste de tamaño de la ventana
root.geometry("1000x300")
# Título del proyecto
root.title("Proyecto ISP")
root.configure(background="#ADE7FF")
# Cambio de icono
BASE_DIR = os.path.dirname((os.path.abspath(__file__)))
ruta = os.path.join(BASE_DIR, "img", "icono.png")
image_root = Image.open(ruta)
image_icon = ImageTk.PhotoImage(image_root)
root.iconphoto(True, image_icon)

############
# --Menu-- #
############
menubar = Menu(root)

menu_archivo = Menu(menubar, tearoff=0)
menu_archivo.add_command(label="Abrir", command=abrir_archivo)
menu_archivo.add_command(label="Guardar como...", command=guardar_archivo)
menu_archivo.add_separator()
menu_archivo.add_command(label="Salir", command=root.quit)
menubar.add_cascade(label="Archivo", menu=menu_archivo)

root.config(menu=menubar)

###############
# --Botones-- #
###############
# Botón agregar cliente
b_agregar_cliente = Button(
    root, text="Agregar cliente", command=f_agregar_cliente, width=15
)
b_agregar_cliente.grid(row=0, column=0, sticky=W, padx=20, pady=10)
# grid_propagate expande las grillas en conjunto con la ventana principal
b_agregar_cliente.grid_propagate(True)
b_agregar_cliente.configure(border=3)
# Botón consulta
b_consulta = Button(root, text="Consulta", command=f_consulta, width=15)
b_consulta.grid(row=0, column=1, sticky=W, padx=20, pady=10)
b_consulta.grid_propagate(True)
b_consulta.configure(border=3)
# Botón modificación
b_modificacion = Button(root, text="Modificación", command=f_modificacion, width=15)
b_modificacion.grid(row=0, column=2, sticky=W, padx=20, pady=10)
b_modificacion.grid_propagate(True)
b_modificacion.configure(border=3)
# Botón baja
b_baja = Button(root, text="Baja", command=f_baja, width=15)
b_baja.grid(row=0, column=3, sticky=W, padx=20, pady=10)
b_baja.grid_propagate(True)
b_baja.configure(border=3)
# Botón exportar
b_exportar = Button(root, text="Export", command=f_exportar, width=15)
b_exportar.grid(row=0, column=6, sticky=W, padx=20, pady=10)
b_exportar.grid_propagate(True)
b_exportar.configure(border=3)

#############
# --Árbol-- #
#############
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
tree.grid(row=3, column=0, columnspan=7, padx=20, pady=10)

root.mainloop()
