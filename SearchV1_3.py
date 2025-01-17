import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pyodbc

# Conexión a la base de datos SQL Server
def conectar_bd():
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=sistemas01\\SQLEXPRESS;'
        'DATABASE=InvenIT;'
        'Trusted_Connection=yes'
    )
    return conn

# Mapeo entre las cadenas descriptivas y los nombres de las columnas
criterios_mapeo = {
    "Puesto de Trabajo": "W.nombre",
    "Nombre de Usuario": "U.username",
    "Nombre de PC": "P.nombre",
    "Código de PC": "P.codigo",
    "Código de Monitor": "M.codigo"
}

# Función para buscar datos en la base de datos
def buscar_datos():
    busqueda = busqueda_entry.get()
    criterio_descriptivo = criterio_var.get()
    criterio_real = criterios_mapeo[criterio_descriptivo]

    # Consultar la base de datos
    conn = conectar_bd()
    cursor = conn.cursor()
    query = f"""
        SELECT W.nombre, U.username, P.nombre, P.codigo, M.codigo
        FROM WORKSTAT W
        INNER JOIN USUARIO U ON W.ID_WORKSTAT = U.ID_WORKSTAT
        INNER JOIN PC P ON W.ID_WORKSTAT = P.ID_WORKSTAT
        INNER JOIN MONITOR M ON P.ID_PC = M.ID_PC
        WHERE {criterio_real} LIKE ?
    """
    cursor.execute(query, ('%' + busqueda + '%',))
    resultados = cursor.fetchall()

    # Limpiar la tabla de resultados antes de agregar nuevos
    for row in tabla.get_children():
        tabla.delete(row)

    # Mostrar los resultados en la tabla
    for row in resultados:
        tabla.insert("", "end", values=list(row))

    conn.close()

# Crear la ventana principal
root = ttk.Window(themename="superhero")
root.title("FABEN S.A. - César Barbero")
root.geometry("1100x500")
root.iconbitmap("BuscaIT.ico")

# Encabezado
header = ttk.Label(root, text="Rastreo de activos informáticos", font=("Helvetica", 18, "bold"), bootstyle="primary")
header.pack(pady=20)

# Contenedor principal
frame = ttk.Frame(root, padding=20)
frame.pack(fill=BOTH, expand=True)

# Seleccionar el criterio de búsqueda
criterio_label = ttk.Label(frame, text="Buscar por:", font=("Helvetica", 12))
criterio_label.grid(row=0, column=0, padx=10, pady=10, sticky=W)

criterio_var = ttk.StringVar(value="Puesto de Trabajo")
criterio_menu = ttk.Combobox(frame, textvariable=criterio_var, values=list(criterios_mapeo.keys()), bootstyle="info")
criterio_menu.grid(row=0, column=1, padx=10, pady=10, sticky=W)

# Campo de entrada para la búsqueda
busqueda_label = ttk.Label(frame, text="Ingrese dato a buscar:", font=("Helvetica", 12))
busqueda_label.grid(row=1, column=0, padx=10, pady=10, sticky=W)

busqueda_entry = ttk.Entry(frame, font=("Helvetica", 12), bootstyle="success")
busqueda_entry.grid(row=1, column=1, padx=10, pady=10, sticky=W)

# Botón para realizar la búsqueda
buscar_button = ttk.Button(frame, text="Buscar", command=buscar_datos, bootstyle="primary")
buscar_button.grid(row=2, column=0, columnspan=2, pady=20)

# Tabla para mostrar los resultados
tabla_frame = ttk.Frame(root, padding=10)
tabla_frame.pack(fill=BOTH, expand=True)

# Configuración de estilo para aumentar el tamaño de los datos y las filas
style = ttk.Style()
style.configure("Treeview", rowheight=40, font=("Helvetica", 9))  # Aumentar el tamaño de las filas y la fuente
style.configure("Treeview.Heading", font=("Helvetica", 11, "bold"), anchor="center")
style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

# Configurar la tabla con columnas
tabla = ttk.Treeview(tabla_frame, columns=(
    "Puesto de Trabajo", "Nombre de Usuario", "Nombre de PC", "Código de PC", "Código de Monitor"),
    show="headings", height=10, bootstyle="light"
)

tabla.heading("Puesto de Trabajo", text="Puesto de Trabajo", anchor="center")
tabla.heading("Nombre de Usuario", text="Nombre de Usuario", anchor="center")
tabla.heading("Nombre de PC", text="Nombre de PC", anchor="center")
tabla.heading("Código de PC", text="Código de PC", anchor="center")
tabla.heading("Código de Monitor", text="Código de Monitor", anchor="center")

tabla.column("Puesto de Trabajo", anchor="center", width=200)
tabla.column("Nombre de Usuario", anchor="center", width=150)
tabla.column("Nombre de PC", anchor="center", width=150)
tabla.column("Código de PC", anchor="center", width=150)
tabla.column("Código de Monitor", anchor="center", width=150)

tabla.pack(fill=BOTH, expand=True, padx=20, pady=10)

# Barra de desplazamiento
scrollbar = ttk.Scrollbar(tabla_frame, orient="vertical", command=tabla.yview)
scrollbar.pack(side=RIGHT, fill=Y)
tabla.configure(yscrollcommand=scrollbar.set)

autor_label = ttk.Label(root, text="Dev: César Barbero", font=("Helvetica", 12, "italic"))
autor_label.pack(side="bottom", pady=20)


# Ejecutar la interfaz gráfica
root.mainloop()
