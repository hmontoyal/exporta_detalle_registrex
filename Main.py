import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk  # Pillow para manejar el logo

# === Clase para seleccionar cliente ===
class SelectorClientes(tk.Toplevel):
    def __init__(self, root, clientes):
        super().__init__(root)
        self.title("Selecciona un cliente")
        self.geometry("400x300")
        self.clientes = clientes
        self.selected_id = None

        tk.Label(self, text="Selecciona un cliente:").pack(pady=10)

        self.listbox = tk.Listbox(self, width=50, height=12)
        for cliente in clientes:
            self.listbox.insert(tk.END, cliente[1])
        self.listbox.pack()

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Seleccionar", command=self.seleccionar).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Cancelar", command=self.cancelar).pack(side=tk.LEFT, padx=5)

    def seleccionar(self):
        idx = self.listbox.curselection()
        if not idx:
            messagebox.showwarning("Atención", "Debes seleccionar un cliente.")
            return
        idx = idx[0]
        self.selected_id = self.clientes[idx][0]
        self.destroy()

    def cancelar(self):
        self.selected_id = None
        self.destroy()

# === Función principal ===
def main():
    root = tk.Tk()
    root.withdraw()

    # Mostrar ventana de bienvenida con logo y botón
    bienvenida = tk.Toplevel()
    bienvenida.title("Bienvenido")
    bienvenida.geometry("400x300")
    bienvenida.resizable(False, False)

    # Cargar imagen del logo
    try:
        logo_path = "logo.png"  # Asegúrate que el archivo esté en el mismo directorio
        logo_img = Image.open(logo_path)
        logo_img = logo_img.resize((200, 50), Image.ANTIALIAS)
        logo_tk = ImageTk.PhotoImage(logo_img)

        label_logo = tk.Label(bienvenida, image=logo_tk)
        label_logo.image = logo_tk  # Guardar referencia
        label_logo.pack(pady=10)
    except Exception as e:
        print("⚠️ No se pudo cargar el logo:", e)

    tk.Label(bienvenida, text="Presiona el botón para seleccionar un archivo Excel").pack(pady=10)

    archivo_seleccionado = {"path": None}

    def seleccionar_archivo():
        path = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if path:
            archivo_seleccionado["path"] = path
            bienvenida.destroy()
        else:
            messagebox.showinfo("Cancelado", "No se seleccionó ningún archivo.")

    tk.Button(bienvenida, text="Seleccionar archivo", command=seleccionar_archivo).pack(pady=10)
    root.wait_window(bienvenida)

    archivo_origen = archivo_seleccionado["path"]
    if not archivo_origen:
        return

    try:
        df_clientes = pd.read_excel(archivo_origen, sheet_name="DatosCliente")
        df_clientes = df_clientes.sort_values(by="Id", ascending=False)
        clientes = list(zip(df_clientes["Id"], df_clientes["Razon Social"]))

        selector = SelectorClientes(root, clientes)
        root.wait_window(selector)

        id_cliente = selector.selected_id
        if id_cliente is None:
            messagebox.showinfo("Cancelado", "No se seleccionó ningún cliente.")
            return

        hoja = "DatosEquipamiento"
        columnas_a_exportar = [
            "Marca", "Modelo", "N Serie", "N Serie Bandeja",
            "Mueble", "IP", "KFS/SDS", "Ubicación", "Piso",
            "Observacion", "Fecha"
        ]

        df_equip = pd.read_excel(archivo_origen, sheet_name=hoja)
        df_filtrado = df_equip[df_equip["IdRef"] == id_cliente]

        if df_filtrado.empty:
            messagebox.showwarning("Sin resultados", "No se encontraron registros para el cliente seleccionado.")
            return

        df_exportado = df_filtrado[columnas_a_exportar]

        archivo_destino = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Guardar archivo como",
            initialfile="Equipamiento_filtrado.xlsx"
        )

        if archivo_destino:
            df_exportado.to_excel(archivo_destino, index=False)
            messagebox.showinfo("Éxito", f"Archivo exportado como:\n{archivo_destino}")
        else:
            messagebox.showinfo("Cancelado", "Exportación cancelada.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

# === Ejecutar ===
if __name__ == "__main__":
    main()
