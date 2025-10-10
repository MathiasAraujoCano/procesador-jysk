import os
import tkinter as tk
from tkinter import filedialog, messagebox
from procesar_lotes_core import procesar_lotes

class DragDropApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Procesador JYSK - Resumen de Lotes")
        self.geometry("500x350")
        self.configure(bg="#565c83")
        self.resizable(False, False)

        self.instruccion = tk.Label(self, text="Selecciona el archivo que yo hago el trabajo ☕", font=("Arial", 14), bg="#565c83", fg="white")
        self.instruccion.pack(pady=40)

        self.select_btn = tk.Button(self, text="Seleccionar archivo", command=self.select_file, font=("Arial", 14), bg="#f7f7f7", fg="#565c83", activebackground="#e0e0e0", activeforeground="#565c83", relief="solid", borderwidth=2, highlightthickness=0, width=20, height=2)
        self.select_btn.pack(pady=20)
        self.select_btn.bind("<Enter>", self.on_hover)
        self.select_btn.bind("<Leave>", self.on_leave)

        self.status = tk.Label(self, text="", font=("Arial", 12), bg="#565c83", fg="white")
        self.status.pack(pady=10)

    def on_hover(self, event):
        self.select_btn.config(bg="#e0e0e0", cursor="hand2")

    def on_leave(self, event):
        self.select_btn.config(bg="#f7f7f7", cursor="arrow")

    def select_file(self, event=None):
        self.select_btn.config(state="disabled", text="Procesando... ⏳")
        self.status.config(text="Cargando...", fg="white")
        self.update()
        file_path = filedialog.askopenfilename(
            title="Selecciona el archivo en formato xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.process_file(file_path)
            self.select_btn.config(state="normal", text="Seleccionar archivo")
        else:
            self.status.config(text="", fg="white")
            self.select_btn.config(state="normal", text="Seleccionar archivo")

    def process_file(self, file_path):
        self.status.config(text="Procesando archivo...", fg="#333")
        self.update()
        try:
            resultado = procesar_lotes(file_path)
            output_file = resultado['output_file']
            resumen = resultado['resumen']
            fechas_unicas = resumen['Fecha'].unique()
            msg = f"✅ Proceso terminado.\nArchivo generado: {output_file}\n\n"
            msg += "Se crearon las siguientes hojas por fecha:\n"
            for fecha in sorted(fechas_unicas):
                cantidad = len(resumen[resumen['Fecha'] == fecha])
                try:
                    if hasattr(fecha, 'strftime'):
                        fecha_display = fecha.strftime("%d-%m-%Y")
                    else:
                        fecha_display = str(fecha).replace('/', '-').replace(' 00:00:00', '')[:10]
                except:
                    fecha_display = str(fecha)[:10]
                msg += f"  • Fecha {fecha_display}: {cantidad} registros\n"
            msg += f"\n💰 Monto total procesado: ${resumen['Monto_Total'].sum():,.0f}\n"
            msg += f"🔢 Total de transacciones: {resumen['Cantidad_Transacciones'].sum()}\n"
            msg += f"🏪 Cajas procesadas: {len(resumen['Caja_Nombre'].unique())}\n"
            msg += f"💳 Tipos de tarjeta: {', '.join(sorted(resumen['Tarjeta'].unique()))}"
            self.status.config(text="✅ Proceso terminado. Revisa el archivo generado.", fg="white")
            messagebox.showinfo("Resumen generado", msg)
        except Exception as e:
            self.status.config(text=f"❌ Error: {e}", fg="white")
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

def main():
    app = DragDropApp()
    app.mainloop()

if __name__ == "__main__":
    main()
