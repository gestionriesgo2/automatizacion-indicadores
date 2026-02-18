# ======================================================
# INTERFAZ TKINTER CON PROGRESO VISUAL Y SPINNER DE PUNTOS
# ======================================================

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sys
import itertools
import time

# ======================================================
# REDIRIGIR PRINTS A LA INTERFAZ
# ======================================================
class RedirectText:
    def __init__(self, text_widget):
        self.output = text_widget

    def write(self, string):
        self.output.insert(tk.END, string + "\n")
        self.output.see(tk.END)

    def flush(self):
        pass

# ======================================================
# MAIN REAL (TU LÓGICA COMPLETA) CON LOG DE CARPETAS Y ARCHIVOS
# ======================================================
def main():
    # 👇 IMPORTS PESADOS AQUÍ (NO ARRIBA)
    import io
    import pandas as pd
    from datetime import datetime

    from drive_reader import (
        list_files_in_folder,
        read_excel_from_drive,
        get_file_id_by_name,
        create_or_update_file
    )

    from components.banco_drive import (
        cargar_banco_drive,
        clean_str,
        norm_code,
        FICHAS_FOLDER_ID,
        BANCO_FOLDER_ID,
        REPORTE_FOLDER_ID,
        cargar_datos_manuales,
        unir_datos_manuales
    )

    from components.procesar_fichas import procesar_fichas_drive
    from components.guardar_banco_drive import guardar_banco_con_estilos_drive
    from components.guardar_reportes_drive import guardar_reportes_drive

    # ======================================================
    # FECHA AUTOMÁTICA
    # ======================================================
    MESES_ES = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
    }

    HOY = datetime.now()
    ANIO_ACTUAL = str(HOY.year)
    MES_NOMBRE = MESES_ES[HOY.month]
    FECHA_STR = HOY.strftime("%Y-%m-%d")

    COL_CODIGO = "CONSE"
    BANCO_BASE_FILENAME = "Banco_Indicadores_BASE.xlsx"

    print("🚀 Iniciando proceso automático\n")

    # ==================================================
    # OBTENER CARPETAS
    # ==================================================
    def get_or_last_folder(parent_id, preferred_name, label="carpeta"):
        folders = [
            f for f in list_files_in_folder(parent_id)
            if f.get("mimeType") == "application/vnd.google-apps.folder"
        ]
        if not folders:
            raise ValueError(f"No existen {label}s")
        print(f"📁 Carpetas encontradas en '{label}': {[f['name'] for f in folders]}")
        for f in folders:
            if f["name"].lower() == preferred_name.lower():
                print(f"➡ Se selecciona carpeta preferida: {f['name']}")
                return f["id"]
        folders.sort(key=lambda x: x["name"])
        print(f"➡ Se selecciona última carpeta ordenada: {folders[-1]['name']}")
        return folders[-1]["id"]

    FICHAS_ANIO_FOLDER_ID = get_or_last_folder(FICHAS_FOLDER_ID, ANIO_ACTUAL)
    BANCO_ANIO_FOLDER_ID = get_or_last_folder(BANCO_FOLDER_ID, ANIO_ACTUAL)
    BANCO_MES_FOLDER_ID = get_or_last_folder(BANCO_ANIO_FOLDER_ID, MES_NOMBRE)

    # ==================================================
    # CARGAR BANCO
    # ==================================================
    print("🔄 Cargando banco desde Drive...")
    banco, _ = cargar_banco_drive(
        get_file_id_by_name=get_file_id_by_name,
        read_excel_from_drive=read_excel_from_drive
    )
    print(f"📊 Banco cargado con {len(banco)} registros\n")

    # ==================================================
    # BUSCAR FICHAS
    # ==================================================
    files_anio = []
    areas = [
        f for f in list_files_in_folder(FICHAS_FOLDER_ID)
        if f.get("mimeType") == "application/vnd.google-apps.folder"
    ]
    print(f"📁 Áreas encontradas: {[a['name'] for a in areas]}")

    for area in areas:
        nombre_area = area["name"]
        print(f"\n🔹 Leyendo área: {nombre_area}")
        subfolders = [
            f for f in list_files_in_folder(area["id"])
            if f.get("mimeType") == "application/vnd.google-apps.folder"
        ]
        print(f"  📂 Subcarpetas: {[f['name'] for f in subfolders]}")

        carpeta_anio = next((f for f in subfolders if f["name"] == ANIO_ACTUAL), None)
        if not carpeta_anio:
            print(f"  ⚠ No existe carpeta para el año {ANIO_ACTUAL} en {nombre_area}")
            continue

        archivos = [
            f for f in list_files_in_folder(carpeta_anio["id"])
            if f["name"].lower().endswith((".xlsx", ".xlsm"))
        ]
        print(f"  📄 Archivos encontrados: {[f['name'] for f in archivos]}")

        for archivo in archivos:
            files_anio.append({
                "id": archivo["id"],
                "name": archivo["name"],
                "area": nombre_area,
                "anio": ANIO_ACTUAL
            })

    print("\n📊 TOTAL GENERAL DE FICHAS:", len(files_anio))

    # ==================================================
    # PROCESAR FICHAS
    # ==================================================
    print("\n🔄 Procesando fichas...")
    banco, registros = procesar_fichas_drive(
        files_anio=files_anio,
        banco=banco,
        col_codigo=COL_CODIGO,
        read_excel_from_drive=read_excel_from_drive,
        clean_str=clean_str,
        norm_code=norm_code
    )

    print("🔄 Cargando datos manuales...")
    manual_df = cargar_datos_manuales(read_excel_from_drive)
    banco = unir_datos_manuales(banco, manual_df)
    for col in banco.columns:
        banco[col] = banco[col].apply(lambda x: None if isinstance(x, str) and x.startswith("=") else x)

    # ==================================================
    # GUARDAR
    # ==================================================
    print("\n💾 Guardando banco y reportes en Drive...")
    banco_base_id = get_file_id_by_name(BANCO_MES_FOLDER_ID, BANCO_BASE_FILENAME)

    guardar_banco_con_estilos_drive(
        banco=banco,
        create_or_update_file=create_or_update_file,
        banco_file_id=banco_base_id,
        banco_folder_id=BANCO_MES_FOLDER_ID,
        filename=BANCO_BASE_FILENAME
    )

    guardar_banco_con_estilos_drive(
        banco=banco,
        create_or_update_file=create_or_update_file,
        banco_file_id=None,
        banco_folder_id=BANCO_MES_FOLDER_ID,
        filename=f"Banco_Indicadores_{FECHA_STR}.xlsx"
    )

    guardar_reportes_drive(
        registros=registros,
        get_file_id_by_name=get_file_id_by_name,
        create_or_update_file=create_or_update_file,
        reporte_folder_id=REPORTE_FOLDER_ID
    )

    print("\n🎉 Proceso finalizado correctamente")

# ======================================================
# SPINNER DE PUNTOS ANIMADOS
# ======================================================
spinner_running = False
dots_cycle = itertools.cycle(["", ".", "..", "..."])

def update_spinner():
    if spinner_running:
        boton.config(text=f"Cargando{next(dots_cycle)}")
        ventana.after(500, update_spinner)
    else:
        boton.config(text="Ejecutar Proceso")


# ======================================================
# EJECUTAR PROCESO EN HILO
# ======================================================
def ejecutar_proceso():
    global spinner_running
    try:
        spinner_running = True
        progress.start(10)
        update_spinner()  # iniciar animación de puntos
        main()
    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        spinner_running = False
        progress.stop()
        boton.config(state="normal")
        log_text.insert(tk.END, "\n✅ Proceso completado!\n")
        messagebox.showinfo("Éxito", "Proceso finalizado correctamente")


def iniciar():
    boton.config(state="disabled")
    log_text.delete(1.0, tk.END)
    hilo = threading.Thread(target=ejecutar_proceso)
    hilo.start()


# ======================================================
# ENTRY POINT
# ======================================================
if __name__ == "__main__":

    # 🔹 Si se ejecuta con argumento "auto"
    if len(sys.argv) > 1 and sys.argv[1] == "auto":
        print("🤖 Ejecutando en modo automático (cron)...\n")
        try:
            main()
            print("\n✅ Proceso finalizado correctamente (modo automático)")
        except Exception as e:
            print("\n❌ Error en modo automático:")
            print(str(e))
        sys.exit()

    # 🔹 Si NO tiene argumento → abre interfaz gráfica
    ventana = tk.Tk()
    ventana.title("Sistema Matriz de Indicadores")
    ventana.geometry("650x500")

    titulo = tk.Label(ventana, text="Matriz de Indicadores", font=("Arial", 16, "bold"))
    titulo.pack(pady=10)

    boton = tk.Button(ventana, text="Ejecutar Proceso", command=iniciar, width=20, height=2)
    boton.pack(pady=10)

    progress = ttk.Progressbar(ventana, orient="horizontal", mode="indeterminate", length=400)
    progress.pack(pady=10)

    log_text = tk.Text(ventana, height=15, width=80)
    log_text.pack(pady=10)

    sys.stdout = RedirectText(log_text)

    ventana.mainloop()
