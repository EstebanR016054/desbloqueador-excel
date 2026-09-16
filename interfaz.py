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

        # 5. Aplicar reglas, desproteger y modificar Z3
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

            # NUEVA FUNCIÓN: Modificar celda Z3 (Solo si la hoja ya está desprotegida)
            if not ws.ProtectContents:
                # a) Convertir a texto, quitar espacios a los lados y poner en mayúsculas
                texto_z3 = str(valor_z3_crudo).strip().upper()
                
                # b) Validar si termina en "C". Si no, se la agregamos.
                if not texto_z3.endswith("C"):
                    texto_z3 += "C"
                    
                # c) Escribir el nuevo valor en la celda (esto mantiene intacto el formato visual)
                ws.Range("Z3").Value = texto_z3

        # 6. Guardar y cerrar
        wb.Save()
        wb.Close()
        excel.Quit()

        messagebox.showinfo("Éxito", f"¡Proceso terminado!\nSe completó la validación y limpieza del archivo.")
        estado_label.config(text="Esperando nuevo archivo...")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{e}")
        estado_label.config(text="Error en el proceso.")
        try:
            excel.Quit() # Asegurar que Excel se cierre si hay error
        except:
            pass

# ==========================================
# CONFIGURACIÓN DE LA VENTANA VISUAL
# ==========================================
ventana = tk.Tk()
ventana.title("Desbloqueador de Excel")
ventana.geometry("450x250")
ventana.configure(padx=20, pady=20)

tk.Label(ventana, text="Contraseñas (separadas por coma):", font=("Arial", 10, "bold")).pack(pady=5)

# Caja de texto para las claves (puedes dejar unas por defecto)
entrada_claves = tk.Entry(ventana, width=50)
entrada_claves.insert(0, "clave1, clave2, 12345, admin") 
entrada_claves.pack(pady=5)

tk.Label(ventana, text="Filtra hojas, desprotege, y estandariza la celda Z3.", fg="gray").pack(pady=10)

# Botón principal
btn_procesar = tk.Button(ventana, text="Cargar Archivo y Procesar", command=procesar_excel, bg="#0078D7", fg="white", font=("Arial", 11, "bold"))
btn_procesar.pack(pady=10)

# Etiqueta de estado
estado_label = tk.Label(ventana, text="Esperando archivo...")
estado_label.pack(pady=5)

ventana.mainloop()