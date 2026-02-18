import threading
import tkinter as tk
from tkinter import messagebox
from main import main   # pyright: ignore[reportMissingImports] # 👈 importa tu función principal


def ejecutar_proceso():
    try:
        boton.config(state="disabled")
        estado_label.config(text="⏳ Ejecutando proceso...")

        # Ejecutar en hilo para que no se congele la ventana
        hilo = threading.Thread(target=proceso)
        hilo.start()

    except Exception as e:
        messagebox.showerror("Error", str(e))


def proceso():
    try:
        main()
        estado_label.config(text="✅ Proceso finalizado correctamente")
        messagebox.showinfo("Éxito", "El proceso terminó correctamente")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        estado_label.config(text="❌ Error en el proceso")
    finally:
        boton.config(state="normal")


# =========================
# VENTANA PRINCIPAL
# =========================
ventana = tk.Tk()
ventana.title("Sistema de Indicadores")
ventana.geometry("400x250")
ventana.resizable(False, False)

titulo = tk.Label(
    ventana,
    text="Sistema Automático de Indicadores",
    font=("Arial", 14, "bold")
)
titulo.pack(pady=20)

boton = tk.Button(
    ventana,
    text="Ejecutar Proceso",
    font=("Arial", 12),
    width=20,
    height=2,
    bg="#4CAF50",
    fg="white",
    command=ejecutar_proceso
)
boton.pack(pady=20)

estado_label = tk.Label(
    ventana,
    text="Listo para ejecutar",
    font=("Arial", 10)
)
estado_label.pack(pady=10)

ventana.mainloop()
