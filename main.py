import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import pandas as pd
from datetime import datetime

# Importamos los módulos de plantillas:
from plantillacode2 import generar_documento_desde_plantilla, generar_documento_desde_plantilla2
import plantillacode1    # (Formatos 3 y 4 – COMPIN múltiples)
import plantillacode3    # (Formatos 5 y 6 – ISAPRE simples)
import plantillacode4    # (Formatos 7 y 8 – ISAPRE múltiples)

# --- Variables globales para la interfaz ---
root = None
tree = None
btn_generar = None
btn_cargar = None
plantilla_var = None

# DataFrame con los datos cargados desde el Excel
df_licencias = None


# ---------------------------------------------------
# Diálogo para elegir un archivo Excel/CSV en “base de datos”
# ---------------------------------------------------
def dialogo_elegir_excel():
    """
    Abre un filedialog para elegir un .xlsx o .csv dentro de la carpeta ‘base de datos’.
    Si el usuario cancela, retorna None.
    """
    carpeta_bd = "base de datos"
    if not os.path.exists(carpeta_bd):
        os.makedirs(carpeta_bd)
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        initialdir=carpeta_bd,
        filetypes=[("Archivos Excel", "*.xlsx"), ("CSV", "*.csv"), ("Todos los archivos", "*.*")]
    )
    if not ruta:
        return None
    return ruta


# ---------------------------------------------------
# Diálogo para elegir la ISAPRE (solo Formatos 5–8)
# ---------------------------------------------------
class IsapreDialog(tk.Toplevel):
    """
    Permite escoger entre Banmédica, Cruz Blanca, Vida Tres, Nueva Mas Vida, Colmena y Consalud.
    Retorna en self.resultado el nombre de la ISAPRE elegida.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Seleccione la ISAPRE")
        self.geometry("300x280")
        self.resizable(False, False)
        self.resultado = None

        tk.Label(self, text="Elija la ISAPRE:", font=("Arial", 11, "bold")).pack(pady=10)

        self.isapre_var = tk.StringVar(value="Banmédica")
        opciones = ["Banmédica", "Cruz Blanca", "Vida Tres", "Nueva Mas Vida", "Colmena", "Consalud"]
        for op in opciones:
            ttk.Radiobutton(self, text=op, variable=self.isapre_var, value=op).pack(anchor="w", padx=20, pady=3)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="Aceptar", command=self.on_ok).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Cancelar", command=self.on_cancel).pack(side="left", padx=10)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.wait_window()

    def on_ok(self):
        self.resultado = self.isapre_var.get()
        self.destroy()

    def on_cancel(self):
        self.resultado = None
        self.destroy()


# ---------------------------------------------------
# Diálogos para Datos Extra y Subrogancia
# ---------------------------------------------------
class DatosExtraDialog(tk.Toplevel):
    def __init__(self, parent, licencia=None):
        super().__init__(parent)
        self.title("Datos Extra Decreto")
        self.geometry("400x300")
        self.resizable(False, False)
        self.resultado = None

        # Si “licencia” ya trae campos de Excel, los precargamos
        self.var_da = tk.StringVar(value=licencia["decreto_aut_excel"] if licencia and "decreto_aut_excel" in licencia else "")
        self.var_fecha = tk.StringVar(value=licencia["fecha_decreto_excel"] if licencia and "fecha_decreto_excel" in licencia else "")
        self.var_secretario = tk.StringVar()

        ttk.Label(self, text="N° Decreto Alcaldicio:").pack(anchor="w", padx=20, pady=(15, 0))
        ttk.Entry(self, textvariable=self.var_da).pack(fill="x", padx=20)
        ttk.Label(self, text="Fecha Decreto (DD/MM/AAAA):").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(self, textvariable=self.var_fecha).pack(fill="x", padx=20)
        ttk.Label(self, text="Nombre completo del Secretario Municipal:").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(self, textvariable=self.var_secretario).pack(fill="x", padx=20)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="Aceptar", command=self.on_ok).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Cancelar", command=self.on_cancel).pack(side="left", padx=10)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.wait_window()

    def on_ok(self):
        if (not self.var_da.get().strip()
                or not self.var_fecha.get().strip()
                or not self.var_secretario.get().strip()):
            messagebox.showerror("Error", "Todos los campos son obligatorios.", parent=self)
            return
        self.resultado = {
            "decreto_aut_excel": self.var_da.get().strip(),
            "fecha_decreto_excel": self.var_fecha.get().strip(),
            "secretario": self.var_secretario.get().strip()
        }
        self.destroy()

    def on_cancel(self):
        self.resultado = None
        self.destroy()


class SubroganciaDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Texto de Subrogancia (Viñeta c)")
        self.geometry("500x220")
        self.resizable(False, False)
        self.resultado = None

        ttk.Label(self, text="Ingresa el texto para la viñeta c) de subrogancia:").pack(anchor="w", padx=20, pady=(15, 0))
        self.text = tk.Text(self, height=5, width=60)
        self.text.pack(padx=20, pady=10)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Aceptar", command=self.on_ok).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Cancelar", command=self.on_cancel).pack(side="left", padx=10)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.wait_window()

    def on_ok(self):
        texto = self.text.get("1.0", "end").strip()
        if not texto:
            messagebox.showerror("Error", "Debes ingresar el texto de subrogancia.", parent=self)
            return
        self.resultado = texto
        self.destroy()

    def on_cancel(self):
        self.resultado = None
        self.destroy()


class DatosSubroganciaDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Datos de Subrogancia")
        self.geometry("520x320")
        self.resizable(False, False)
        self.resultado = None

        canvas = tk.Canvas(self, borderwidth=0, width=500, height=250)
        frame_scroll = ttk.Frame(canvas)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=frame_scroll, anchor="nw")

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame_scroll.bind("<Configure>", on_frame_configure)

        self.var_nombre = tk.StringVar()
        self.var_genero = tk.StringVar(value="F")
        self.var_trato = tk.StringVar()
        self.var_cargo = tk.StringVar()
        self.var_direccion = tk.StringVar()
        self.var_decreto = tk.StringVar()
        self.var_fecha_decreto = tk.StringVar()
        self.var_desde = tk.StringVar()
        self.var_hasta = tk.StringVar()

        ttk.Label(frame_scroll, text="Nombre completo de quien subroga:").pack(anchor="w", padx=20, pady=(15, 0))
        ttk.Entry(frame_scroll, textvariable=self.var_nombre).pack(fill="x", padx=20)
        ttk.Label(frame_scroll, text="Género de quien subroga:").pack(anchor="w", padx=20, pady=(10, 0))
        frame_gen = ttk.Frame(frame_scroll)
        frame_gen.pack(anchor="w", padx=20)
        ttk.Radiobutton(frame_gen, text="Femenino", variable=self.var_genero, value="F").pack(side="left", padx=5)
        ttk.Radiobutton(frame_gen, text="Masculino", variable=self.var_genero, value="M").pack(side="left", padx=5)
        ttk.Label(frame_scroll, text="Trato de quien subroga (Sr./Sra.):").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(frame_scroll, textvariable=self.var_trato).pack(fill="x", padx=20)
        ttk.Label(frame_scroll, text="Cargo de quien subroga:").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(frame_scroll, textvariable=self.var_cargo).pack(fill="x", padx=20)
        ttk.Label(frame_scroll, text="Dirección a subrogar:").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(frame_scroll, textvariable=self.var_direccion).pack(fill="x", padx=20)
        ttk.Label(frame_scroll, text="N° Decreto de Subrogancia:").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(frame_scroll, textvariable=self.var_decreto).pack(fill="x", padx=20)
        ttk.Label(frame_scroll, text="Fecha Decreto Subrogancia (DD/MM/AAAA):").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(frame_scroll, textvariable=self.var_fecha_decreto).pack(fill="x", padx=20)
        ttk.Label(frame_scroll, text="Desde (DD/MM/AAAA):").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(frame_scroll, textvariable=self.var_desde).pack(fill="x", padx=20)
        ttk.Label(frame_scroll, text="Hasta (DD/MM/AAAA):").pack(anchor="w", padx=20, pady=(10, 0))
        ttk.Entry(frame_scroll, textvariable=self.var_hasta).pack(fill="x", padx=20)

        btn_frame = ttk.Frame(frame_scroll)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="Aceptar", command=self.on_ok).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Cancelar", command=self.on_cancel).pack(side="left", padx=10)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.wait_window()

        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

    def on_ok(self):
        campos = [
            self.var_nombre.get().strip(),
            self.var_trato.get().strip(),
            self.var_cargo.get().strip(),
            self.var_direccion.get().strip(),
            self.var_decreto.get().strip(),
            self.var_fecha_decreto.get().strip(),
            self.var_desde.get().strip(),
            self.var_hasta.get().strip()
        ]
        if any(not c for c in campos):
            messagebox.showerror("Error", "Todos los campos son obligatorios.", parent=self)
            return

        self.resultado = {
            "nombre_subrogante": self.var_nombre.get().strip(),
            "genero_subrogante": self.var_genero.get().strip(),
            "trato_subrogante": self.var_trato.get().strip(),
            "cargo_subrogante": self.var_cargo.get().strip(),
            "direccion_subrogada": self.var_direccion.get().strip(),
            "decreto_subrogancia": self.var_decreto.get().strip(),
            "fecha_decreto_subrogancia": self.var_fecha_decreto.get().strip(),
            "desde_subrogancia": self.var_desde.get().strip(),
            "hasta_subrogancia": self.var_hasta.get().strip()
        }
        self.destroy()

    def on_cancel(self):
        self.resultado = None
        self.destroy()


# ---------------------------------------------------
# Funciones auxiliares y de “Cargando…”
# ---------------------------------------------------
def mostrar_cargando():
    global ventana_cargando, root
    if root is None:
        return
    ventana_cargando = tk.Toplevel(root)
    ventana_cargando.title("Generando documento…")
    ventana_cargando.geometry("300x80")
    ventana_cargando.resizable(False, False)
    tk.Label(ventana_cargando, text="Generando documento, por favor espera…").pack(pady=20)
    ventana_cargando.transient(root)
    ventana_cargando.grab_set()
    ventana_cargando.protocol("WM_DELETE_WINDOW", lambda: None)


def cerrar_cargando():
    global ventana_cargando
    if ventana_cargando is not None:
        ventana_cargando.destroy()
        ventana_cargando = None


def limpiar_valor_excel_general(valor):
    """
    Limpia valores de Excel (None, nan, etc.) a cadena vacía si corresponde.
    """
    if pd.isnull(valor):
        return ""
    return str(valor).strip()


def preguntar_genero(nombre):
    """
    Si el Excel no tiene género, abre este diálogo para elegir F o M.
    """
    global root
    win = tk.Toplevel(root)
    win.title("Selecciona el género")
    win.geometry("350x120")
    win.resizable(False, False)
    tk.Label(win, text=f"Selecciona el género de:\n{nombre}", font=("Arial", 11)).pack(pady=10)
    genero_var = tk.StringVar(value="F")
    frame = ttk.Frame(win)
    frame.pack()
    ttk.Radiobutton(frame, text="Femenino", variable=genero_var, value="F").pack(side="left", padx=10)
    ttk.Radiobutton(frame, text="Masculino", variable=genero_var, value="M").pack(side="left", padx=10)
    confirmado = tk.BooleanVar(value=False)

    def confirmar():
        confirmado.set(True)
        win.destroy()

    ttk.Button(win, text="Aceptar", command=confirmar).pack(pady=10)
    win.transient(root)
    win.grab_set()
    win.wait_variable(confirmado)
    return genero_var.get()


# ---------------------------------------------------
# Al pulsar “Cargar Excel”: cargar datos y mostrar en el Treeview
# ---------------------------------------------------
def cargar_y_mostrar():
    """
    Llama a cargar_datos_excel() para obtener un DataFrame válido.
    Si lo obtiene, limpia el Treeview y lo rellena con las filas.
    Habilita el botón “Generar Decreto” y resetea la selección previa.
    """
    global df_licencias, tree, btn_generar

    df = cargar_datos_excel()
    if df is None:
        return

    df_licencias = df

    # Limpiamos cualquier fila previa en el Treeview
    for fila in tree.get_children():
        tree.delete(fila)

    # Insertamos cada fila del DataFrame en el Treeview
    for _, row in df.iterrows():
        tree.insert(
            "",
            "end",
            values=(
                row.get("N° De Licencia", ""),
                row.get("NOMBRE_APELLIDOS", ""),
                row.get("RUT", ""),
                row.get("ESCALAFON", ""),
                row.get("GRADO", ""),
                row.get("Desde", ""),
                row.get("Hasta", ""),
                row.get("DECRETO ALCALDICIO", ""),
                row.get("FECHA DE DECRETO", ""),
                row.get("GÉNERO", "")
            )
        )

    # Habilitamos el botón “Generar Decreto” ahora que hay datos
    btn_generar.config(state=tk.NORMAL)

    # Deseleccionamos todo (por si había algo)
    tree.selection_remove(tree.selection())


# ---------------------------------------------------
# Carga de Excel (invocada por cargar_y_mostrar)
# ---------------------------------------------------
def cargar_datos_excel():
    """
    Abre diálogo para elegir un Excel/CSV de “base de datos”.
    Verifica las columnas mínimas y agrega “GÉNERO” si no existe.
    Retorna el DataFrame o None si hubo cancelación/ error.
    """
    ruta = dialogo_elegir_excel()
    if not ruta:
        return None

    try:
        ext = os.path.splitext(ruta)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(ruta)
        else:
            df = pd.read_excel(ruta)
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al cargar el archivo:\n{e}")
        return None

    columnas_requeridas = [
        "N° De Licencia", "NOMBRE_APELLIDOS", "RUT", "ESCALAFON",
        "GRADO", "Desde", "Hasta", "DECRETO ALCALDICIO", "FECHA DE DECRETO"
    ]
    for col in columnas_requeridas:
        if col not in df.columns:
            messagebox.showerror("Error", f"Falta la columna '{col}' en el archivo Excel.")
            return None

    # Si falta GÉNERO, la creamos vacía
    if "GÉNERO" not in df.columns:
        df["GÉNERO"] = ""

    # Eliminamos filas sin “N° De Licencia”
    df = df.dropna(subset=["N° De Licencia"])
    return df


# ---------------------------------------------------
# Hilos que generan documentos
# ---------------------------------------------------
def task_generar_documento_1_4(licencia, datos_extra):
    global btn_generar
    try:
        fmt = plantilla_var.get()
        if fmt == "formato1":
            ruta = generar_documento_desde_plantilla(licencia, datos_extra)
        elif fmt == "formato2":
            datos_sub = datos_extra["datos_subrogancia"]
            ruta = generar_documento_desde_plantilla2(licencia, datos_extra, datos_sub)
        elif fmt == "formato3":
            ruta = plantillacode1.generar_documento_formato3([licencia], datos_extra)
        else:  # formato4
            datos_sub = datos_extra["datos_subrogancia"]
            ruta = plantillacode1.generar_documento_formato4([licencia], datos_extra, datos_sub)

        root.after(0, lambda ruta_out=ruta: (
            cerrar_cargando(),
            messagebox.showinfo("¡Listo!", f"Documento generado correctamente:\n\n{ruta_out}"),
            tree.selection_remove(tree.selection()),
            btn_generar.config(state=tk.NORMAL)
        ))

    except Exception as err:
        root.after(0, lambda err=err: (
            cerrar_cargando(),
            messagebox.showerror("Error", f"Ocurrió un error al generar el documento:\n{err}"),
            btn_generar.config(state=tk.NORMAL)
        ))


def task_generar_documento_5_8(lista_licencias, datos_extra, datos_subrogancia, isapre):
    global btn_generar
    try:
        fmt = plantilla_var.get()
        if fmt == "formato5":
            # Solo ISAPRE sin subrogancia
            ruta = plantillacode3.generar_documento_formato5(lista_licencias[0], datos_extra, isapre)
        elif fmt == "formato6":
            # ISAPRE con subrogancia
            datos_sub = datos_subrogancia
            ruta = plantillacode3.generar_documento_formato6(lista_licencias[0], datos_extra, datos_sub, isapre)
        elif fmt == "formato7":
            # Múltiple ISAPRE sin subrogancia
            ruta = plantillacode4.generar_documento_formato7(lista_licencias, datos_extra, isapre)
        else:  # formato8
            ruta = plantillacode4.generar_documento_formato8(lista_licencias, datos_extra, datos_subrogancia, isapre)

        root.after(0, lambda ruta_out=ruta: (
            cerrar_cargando(),
            messagebox.showinfo("¡Listo!", f"Documento generado correctamente:\n\n{ruta_out}"),
            tree.selection_remove(tree.selection()),
            btn_generar.config(state=tk.NORMAL)
        ))

    except Exception as err:
        root.after(0, lambda err=err: (
            cerrar_cargando(),
            messagebox.showerror("Error en Formato 5–8", f"{err}"),
            btn_generar.config(state=tk.NORMAL)
        ))


# ---------------------------------------------------
# Al hacer clic en “Generar Decreto”
# ---------------------------------------------------
def on_generar_click():
    global df_licencias, tree, plantilla_var, btn_generar

    seleccion = tree.selection()
    if not seleccion:
        messagebox.showwarning("Atención", "Selecciona al menos una licencia.")
        return

    # Construir lista de licencias desde las filas seleccionadas
    lista_licencias = []
    for item in seleccion:
        vals = tree.item(item, "values")
        lic = {
            "id": vals[0],
            "nombre_titular": vals[1],
            "rut_titular": vals[2],
            "escalafon": vals[3],
            "grado_raw": vals[4],
            "periodo_inicio": vals[5],
            "periodo_fin": vals[6],
            "decreto_aut_excel": limpiar_valor_excel_general(vals[7]),
            "fecha_decreto_excel": limpiar_valor_excel_general(vals[8]),
            "genero": vals[9]
        }
        # Si falta género, preguntamos
        if not lic["genero"] or lic["genero"].strip() == "":
            lic["genero"] = preguntar_genero(lic["nombre_titular"])
        # Calcular días
        try:
            dt_ini = datetime.strptime(lic["periodo_inicio"], "%d/%m/%Y")
            dt_fin = datetime.strptime(lic["periodo_fin"], "%d/%m/%Y")
            lic["dias"] = (dt_fin - dt_ini).days + 1
        except:
            lic["dias"] = ""
        lista_licencias.append(lic)

    # Pedimos DatosExtra (común a todos los formatos)
    dialog = DatosExtraDialog(root, lista_licencias[0])
    datos_extra = dialog.resultado
    if not datos_extra:
        return

    formato = plantilla_var.get()
    datos_subrogancia = None

    # Si el formato (2,4,6,8) requiere subrogancia → abrimos el diálogo de subrogancia
    if formato in ("formato2", "formato4", "formato6", "formato8"):
        dialog_sub = DatosSubroganciaDialog(root)
        datos_subrogancia = dialog_sub.resultado
        if not datos_subrogancia:
            return

    # Si el formato es 5–8 (ISAPRE), pedimos primero la ISAPRE
    if formato in ("formato5", "formato6", "formato7", "formato8"):
        isapre_dialog = IsapreDialog(root)
        isapre_elegida = isapre_dialog.resultado
        if not isapre_elegida:
            return

        # Para formatos 7 y 8 necesitamos pedir “N° de Resolución” para CADA licencia
        resoluciones = {}
        if formato in ("formato7", "formato8"):
            for lic in lista_licencias:
                try:
                    id_lic = str(int(float(lic["id"])))
                except:
                    id_lic = str(lic["id"])
                prompt = f"Licencia Nº {id_lic} - {lic['nombre_titular']}:\nIngrese N° de Resolución {isapre_elegida}:"
                res = simpledialog.askstring("N° de Resolución (ISAPRE)", prompt, parent=root)
                if res is None:
                    return
                resoluciones[id_lic] = res.strip()
            datos_extra["resoluciones"] = resoluciones

        # Lanzamos hilo para Formato 5–8
        btn_generar.config(state=tk.DISABLED)
        mostrar_cargando()
        hilo = threading.Thread(
            target=task_generar_documento_5_8,
            args=(lista_licencias, datos_extra, datos_subrogancia, isapre_elegida),
            daemon=True
        )
        hilo.start()

    else:
        # Formatos 1–4 (COMPIN)
        btn_generar.config(state=tk.DISABLED)
        mostrar_cargando()
        if datos_subrogancia:
            datos_extra["datos_subrogancia"] = datos_subrogancia
        hilo = threading.Thread(
            target=task_generar_documento_1_4,
            args=(lista_licencias[0], datos_extra),
            daemon=True
        )
        hilo.start()


# ---------------------------------------------------
# Inicializar la interfaz principal
# ---------------------------------------------------
def iniciar_interfaz():
    global root, tree, btn_generar, btn_cargar, plantilla_var

    root = tk.Tk()
    root.title("MaxLicenciasApp - Generador de Decretos")
    root.geometry("980x600")

    # Marco contenedor principal
    frame = ttk.Frame(root, padding="10")
    frame.pack(fill=tk.BOTH, expand=True)

    # Botón “Cargar Excel”
    btn_cargar = ttk.Button(root, text="Cargar Excel", command=cargar_y_mostrar)
    btn_cargar.pack(pady=(0, 10))

    # Treeview (vacío al inicio)
    vsb = ttk.Scrollbar(frame, orient="vertical")
    hsb = ttk.Scrollbar(frame, orient="horizontal")
    columns = (
        "id", "nombre_titular", "rut_titular", "escalafon", "grado_raw",
        "desde", "hasta", "decreto_aut_excel", "fecha_decreto_excel", "genero"
    )
    tree = ttk.Treeview(
        frame,
        columns=columns,
        show="headings",
        yscrollcommand=vsb.set,
        xscrollcommand=hsb.set,
        selectmode="extended"
    )
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    encabezados = [
        ("id", "N° De Licencia"),
        ("nombre_titular", "Nombre Titular"),
        ("rut_titular", "RUT Titular"),
        ("escalafon", "Escalafón"),
        ("grado_raw", "Grado"),
        ("desde", "Desde"),
        ("hasta", "Hasta"),
        ("decreto_aut_excel", "D.A. (Excel)"),
        ("fecha_decreto_excel", "Fecha D.A. (Excel)"),
        ("genero", "GÉNERO")
    ]
    for col, text in encabezados:
        tree.heading(col, text=text)
        tree.column(col, width=120 if col == "id" else 140)

    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    hsb.pack(side=tk.BOTTOM, fill=tk.X)
    tree.pack(fill=tk.BOTH, expand=True)

    # Por defecto, el botón “Generar Decreto” está deshabilitado hasta cargar Excel
    plantilla_var = tk.StringVar(value="formato1")
    frame_formato = ttk.LabelFrame(root, text="Seleccione el Formato de Decreto")
    frame_formato.pack(fill="x", padx=20, pady=10)

    # Formatos 1–4 (COMPIN)
    ttk.Radiobutton(frame_formato, text="Formato 1: Solo Compin (sin subrogancia)", variable=plantilla_var, value="formato1").pack(anchor="w", padx=10, pady=2)
    ttk.Radiobutton(frame_formato, text="Formato 2: Solo Compin (con subrogancia)", variable=plantilla_var, value="formato2").pack(anchor="w", padx=10, pady=2)
    ttk.Radiobutton(frame_formato, text="Formato 3: Múltiple Compin (sin subrogancia)", variable=plantilla_var, value="formato3").pack(anchor="w", padx=10, pady=2)
    ttk.Radiobutton(frame_formato, text="Formato 4: Subrogancia Múltiple Compin", variable=plantilla_var, value="formato4").pack(anchor="w", padx=10, pady=2)

    # Formatos 5–8 (ISAPRE)
    ttk.Radiobutton(frame_formato, text="Formato 5: Solo ISAPRE (sin subrogancia)", variable=plantilla_var, value="formato5").pack(anchor="w", padx=10, pady=2)
    ttk.Radiobutton(frame_formato, text="Formato 6: Subrogancia Solo ISAPRE", variable=plantilla_var, value="formato6").pack(anchor="w", padx=10, pady=2)
    ttk.Radiobutton(frame_formato, text="Formato 7: Múltiple ISAPRE (sin subrogancia)", variable=plantilla_var, value="formato7").pack(anchor="w", padx=10, pady=2)
    ttk.Radiobutton(frame_formato, text="Formato 8: Subrogancia Múltiple ISAPRE", variable=plantilla_var, value="formato8").pack(anchor="w", padx=10, pady=2)

    btn_generar = ttk.Button(root, text="Generar Decreto", command=on_generar_click, state=tk.DISABLED)
    btn_generar.pack(pady=15)

    root.mainloop()


if __name__ == "__main__":
    iniciar_interfaz()
