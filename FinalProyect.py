from msilib.schema import ComboBox
from tkinter import *
from tkinter import ttk
from PIL import ImageTk, Image
from tkinter import filedialog as fd
from tkinter.messagebox import *
import os

mi_id = 0

#################
# --Funciones-- #
#################
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


def f_agregar_cliente():
    # Función para pasar datos de la ventana secundaria al treeview de la principal
    def ingreso_datos():
        nonlocal var_cliente, ld_servicios, var_activo
        global tree, mi_id
        nonlocal nueva_ventana
        if ld_servicios.get() != "" and var_cliente != "" and var_activo != "":
            tree.insert(
                "",
                "end",
                text=str(mi_id),
                values=(var_cliente.get(), ld_servicios.get(), var_activo.get()),
            )
            mi_id += 1
            nueva_ventana.destroy()
        else:
            showerror(
                "Rellenar campos de registro",
                "Se detectaron campos vacíos, rellénelos para ingresar el cliente",
            )

    # Variables para ingresar los datos
    var_cliente = StringVar()
    var_servicio = StringVar()
    var_activo = StringVar()

    # Creación de la nueva ventana
    nueva_ventana = Toplevel(root)
    nueva_ventana.title("Agregar Cliente")
    nueva_ventana.geometry("450x230")

    # Labels, entry y botones de la nueva ventana
    l_cliente = Label(nueva_ventana, text="Ingresar cliente:")
    l_cliente.grid(row=0, column=0, padx=10, pady=10, sticky=E)
    cliente = Entry(nueva_ventana, textvariable=var_cliente, width=40)
    cliente.grid(row=0, column=1, padx=10, pady=10)
    l_servicio = Label(nueva_ventana, text="Ingresar servicio:")
    l_servicio.grid(row=1, column=0, padx=10, pady=10, sticky=E)
    # Combobox de servicios a elegir
    valores = ("Servicio1", "Servicio2", "Servicio3", "Servicio4")
    ld_servicios = ttk.Combobox(
        nueva_ventana, width=35, state="readonly", values=valores
    )
    ld_servicios.grid(row=1, column=1, sticky=W, padx=10, pady=10)
    # servicio = Entry(nueva_ventana, textvariable=var_servicio, width=40)
    # servicio.grid(row=1, column=1, padx=10, pady=10)
    l_activo = Label(nueva_ventana, text="Ingresar activo:")
    l_activo.grid(row=2, column=0, padx=10, pady=10, sticky=E)
    activo = Entry(nueva_ventana, textvariable=var_activo, width=40)
    activo.grid(row=2, column=1, padx=10, pady=10)

    b_aceptar_ingreso = Button(nueva_ventana, text="Aceptar", command=ingreso_datos)
    b_aceptar_ingreso.grid(row=3, column=0)
    b_cancelar_ingreso = Button(
        nueva_ventana, text="Cancelar", command=nueva_ventana.destroy
    )
    b_cancelar_ingreso.grid(row=3, column=1)


def f_alta():
    pass


def f_baja():
    pass


def f_servicios():
    pass


def f_modificacion():
    pass


def f_consulta():
    pass


def f_exportar():
    pass


root = Tk()
# Ajuste de tamaño de la ventana
root.geometry("750x500")
# Título del proyecto
root.title("Proyecto ISP")
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
# Botón alta
b_alta = Button(root, text="Alta", command=f_alta, width=15)
b_alta.grid(row=0, column=2, sticky=W, padx=20, pady=10)
b_alta.grid_propagate(True)
b_alta.configure(border=3)
# Botón baja
b_baja = Button(root, text="Baja", command=f_baja, width=15)
b_baja.grid(row=0, column=3, sticky=W, padx=20, pady=10)
b_baja.grid_propagate(True)
b_baja.configure(border=3)
# Botón servicios
b_servicios = Button(root, text="Servicios", command=f_servicios, width=15)
b_servicios.grid(row=1, column=0, sticky=W, padx=20, pady=10)
b_servicios.grid_propagate(True)
b_servicios.configure(border=3)
# Botón modificación
b_modificacion = Button(root, text="Modificación", command=f_modificacion, width=15)
b_modificacion.grid(row=1, column=2, sticky=W, padx=20, pady=10)
b_modificacion.grid_propagate(True)
b_modificacion.configure(border=3)
# Botón consulta
b_consulta = Button(root, text="Consulta", command=f_consulta, width=15)
b_consulta.grid(row=1, column=3, sticky=W, padx=20, pady=10)
b_consulta.grid_propagate(True)
b_consulta.configure(border=3)
# Botón exportar
b_exportar = Button(root, text="Export", command=f_exportar, width=15)
b_exportar.grid(row=2, column=3, sticky=W, padx=20, pady=10)
b_exportar.grid_propagate(True)
b_exportar.configure(border=3)

#############
# --Árbol-- #
#############
tree = ttk.Treeview(root)
tree["columns"] = ("Cliente", "Servicio", "Activo")
# Configuración de las columnas
tree.column("#0", width=150, minwidth=150, anchor=W)
tree.column("Cliente", width=180, minwidth=180, anchor=W)
tree.column("Servicio", width=180, minwidth=180, anchor=W)
tree.column("Activo", width=200, minwidth=200, anchor=W)
# Título de las columnas
tree.heading("#0", text="N° de cliente")
tree.heading("#1", text="Cliente")
tree.heading("#2", text="Servicio")
tree.heading("#3", text="Activo")
tree.grid(row=3, column=0, columnspan=4, padx=20, pady=10)

root.mainloop()
