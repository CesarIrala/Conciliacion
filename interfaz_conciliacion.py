
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from conciliacion_bancaria import generar_excel_conciliacion

def seleccionar_archivo(entry):
    archivo = filedialog.askopenfilename()
    if archivo:
        entry.delete(0, tk.END)
        entry.insert(0, archivo)

def generar_reporte():
    try:
        saldo_texto = entry_saldo.get().strip().replace('.', '').replace(',', '.')
        saldo_inicial = float(saldo_texto)
        extracto = entry_extracto.get()
        vista = entry_vista.get()
        diferidos = entry_diferido.get()
        salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if not all([extracto, vista, diferidos, salida]):
            raise ValueError("Todos los archivos deben estar seleccionados.")

        generar_excel_conciliacion(saldo_inicial, extracto, vista, diferidos, salida)
        messagebox.showinfo("Éxito", f"Archivo generado correctamente:\n{os.path.basename(salida)}")
    except ValueError as ve:
        messagebox.showerror("Error de entrada", str(ve))
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al generar el reporte:\n{str(e)}")

# Interfaz
ventana = tk.Tk()
ventana.title("Generador de Conciliación Bancaria")
ventana.geometry("650x300")
ventana.resizable(False, False)

frame = tk.Frame(ventana)
frame.pack(padx=20, pady=20)

tk.Label(frame, text="Saldo Inicial:").grid(row=0, column=0, sticky="e")
entry_saldo = tk.Entry(frame, width=50)
entry_saldo.grid(row=0, column=1)

tk.Label(frame, text="Extracto Bancario:").grid(row=1, column=0, sticky="e")
entry_extracto = tk.Entry(frame, width=50)
entry_extracto.grid(row=1, column=1)
tk.Button(frame, text="Seleccionar", command=lambda: seleccionar_archivo(entry_extracto)).grid(row=1, column=2)

tk.Label(frame, text="Cheques Emitidos (Vista):").grid(row=2, column=0, sticky="e")
entry_vista = tk.Entry(frame, width=50)
entry_vista.grid(row=2, column=1)
tk.Button(frame, text="Seleccionar", command=lambda: seleccionar_archivo(entry_vista)).grid(row=2, column=2)

tk.Label(frame, text="Cheques Diferidos:").grid(row=3, column=0, sticky="e")
entry_diferido = tk.Entry(frame, width=50)
entry_diferido.grid(row=3, column=1)
tk.Button(frame, text="Seleccionar", command=lambda: seleccionar_archivo(entry_diferido)).grid(row=3, column=2)

boton_generar = tk.Button(ventana, text="Generar Reporte Excel", command=generar_reporte, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
boton_generar.pack(pady=10)

ventana.mainloop()
