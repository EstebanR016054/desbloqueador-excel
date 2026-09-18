import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import pywintypes
import os

def procesar_excel():
    # 1. Abrir ventana para seleccionar el archivo
    ruta_archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
    )
    if not ruta_archivo:
        return # Si el usuario cancela, no hacemos nada

    # 2. Obtener las contraseñas escritas en la interfaz
    contrasenas = entrada_claves.get().split(",")
    contrasenas = [c.strip() for c in contrasenas if c.strip()]

    # 3. Cambiar el texto de estado
    estado_label.config(text="Procesando archivo... Por favor, espera.")
    ventana.update()

    try:
        # 4. Iniciar Excel en segundo plano
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Reemplazar las barras de ruta para evitar errores en Windows
        ruta_archivo = os.path.abspath(ruta_archivo)
        wb = excel.Workbooks.Open(ruta_archivo)
        hojas_modificadas = 0

        # Variable para almacenar el rango B2 de la PRIMERA hoja válida
        rango_b2_origen = None 

        # 5. Aplicar reglas, desproteger, modificar Z3 y copiar/pegar B2 y B33
        for ws in wb.Sheets:
            # Filtro A: Si está oculta, saltar
            if ws.Visible != -1: 
                continue
            
            # Filtro B: Si Z3 está vacía, saltar
            valor_z3_crudo = ws.Range("Z3").Value
            if valor_z3_crudo is None or str(valor_z3_crudo).strip() == "": 
                continue
                
            # Intento de desprotección
            if ws.ProtectContents: 
                for clave in contrasenas:
                    try:
                        ws.Unprotect(Password=clave)
                        hojas_modificadas += 1
                        break # Si funciona, salir del ciclo de claves
                    except pywintypes.com_error:
                        pass # Si falla la clave, intentar la siguiente

            # SOLO SI LA HOJA ESTÁ DESPROTEGIDA Y ES VÁLIDA:
            if not ws.ProtectContents:
                # ----------------------------------------------------
                # ACCIÓN 1: Estandarizar la celda Z3
                # ----------------------------------------------------
                texto_z3 = str(valor_z3_crudo).strip().upper()
                if not texto_z3.endswith("C"):
                    texto_z3 += "C"
                ws.Range("Z3").Value = texto_z3

                # ----------------------------------------------------
                # ACCIÓN 2: Manejo de B2 y B33 (Combinadas + Logo)
                # ----------------------------------------------------
                rango_b2_actual = ws.Range("B2").MergeArea
                rango_b33_actual = ws.Range("B33").MergeArea

                if rango_b2_origen is None:
                    # Es la PRIMERA hoja válida: guardamos la referencia a su rango B2
                    rango_b2_origen = rango_b2_actual
                else:
                    # Es una hoja POSTERIOR: Limpiamos B2 y pegamos
                    for shape in list(ws.Shapes):
                        try:
                            tlc = shape.TopLeftCell
                            if (rango_b2_actual.Row <= tlc.Row <= rango_b2_actual.Row + rango_b2_actual.Rows.Count - 1) and \
                               (rango_b2_actual.Column <= tlc.Column <= rango_b2_actual.Column + rango_b2_actual.Columns.Count - 1):
                                shape.Delete()
                        except:
                            pass
                    rango_b2_actual.ClearContents()
                    rango_b2_origen.Copy(ws.Range("B2"))

                # Ahora, para TODAS las hojas (incluida la primera), limpiamos B33 y pegamos el logo
                for shape in list(ws.Shapes):
                    try:
                        tlc = shape.TopLeftCell
                        # Validar si hay imágenes flotando sobre la zona B33
                        if (rango_b33_actual.Row <= tlc.Row <= rango_b33_actual.Row + rango_b33_actual.Rows.Count - 1) and \
                           (rango_b33_actual.Column <= tlc.Column <= rango_b33_actual.Column + rango_b33_actual.Columns.Count - 1):
                            shape.Delete()
                    except:
                        pass
                
                # Borramos texto/contenido y pegamos manteniendo formato de origen + logo
                rango_b33_actual.ClearContents()
                rango_b2_origen.Copy(ws.Range("B33"))

        # Limpiar el portapapeles de Excel al terminar
        excel.CutCopyMode = False

        # 6. Guardar y cerrar
        wb.Save()
        wb.Close()
        excel.Quit()

        messagebox.showinfo(
            "Éxito", 
            "¡Proceso terminado!\n\n"
            "• Hojas desprotegidas.\n"
            "• Celda Z3 estandarizada.\n"
            "• Logo replicado en B2 y B33 para todas las hojas."
        )
        estado_label.config(text="Esperando nuevo archivo...")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{e}")
        estado_label.config(text="Error en el proceso.")
        try:
            excel.Quit()
        except:
            pass

# ==========================================
# CONFIGURACIÓN DE LA VENTANA VISUAL
# ==========================================
ventana = tk.Tk()
ventana.title("Desbloqueador de Excel")
ventana.geometry("490x270")
ventana.configure(padx=20, pady=20)

tk.Label(ventana, text="Contraseñas (separadas por coma):", font=("Arial", 10, "bold")).pack(pady=5)

# Caja de texto para las claves
entrada_claves = tk.Entry(ventana, width=55)
entrada_claves.insert(0, "clave1, clave2, 12345, admin") 
entrada_claves.pack(pady=5)

tk.Label(
    ventana, 
    text="Filtra hojas, desprotege, estandariza Z3 y replica logo B2 y B33.", 
    fg="gray"
).pack(pady=10)

# Botón principal
btn_procesar = tk.Button(
    ventana, 
    text="Cargar Archivo y Procesar", 
    command=procesar_excel, 
    bg="#0078D7", 
    fg="white", 
    font=("Arial", 11, "bold")
)
btn_procesar.pack(pady=10)

# Etiqueta de estado
estado_label = tk.Label(ventana, text="Esperando archivo...")
estado_label.pack(pady=5)

ventana.mainloop()