import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import os
import subprocess

class ExcelController:
    def __init__(self, master):
        self.master = master
        master.title("Controlador de Excel")

        # Variables
        self.reference = tk.StringVar()
        self.value1 = tk.StringVar()
        self.value2 = tk.StringVar()
        self.value3 = tk.StringVar()
        self.value4 = tk.StringVar()
        self.excel_file = "Salida Equipos.xlsx"  # Nombre por defecto del archivo Excel

        # Crear widgets
        tk.Label(master, text="Equipo:").grid(row=0, column=0, sticky="e")
        self.entry_reference = tk.Entry(master, textvariable=self.reference)
        self.entry_reference.grid(row=0, column=1)

        tk.Label(master, text="SKU:").grid(row=1, column=0, sticky="e")
        self.entry_value1 = tk.Entry(master, textvariable=self.value1, state="disabled")
        self.entry_value1.grid(row=1, column=1)

        tk.Label(master, text="RMA:").grid(row=2, column=0, sticky="e")
        self.entry_value2 = tk.Entry(master, textvariable=self.value2, state="disabled")
        self.entry_value2.grid(row=2, column=1)

        tk.Label(master, text="$:").grid(row=3, column=0, sticky="e")
        self.entry_value3 = tk.Entry(master, textvariable=self.value3, state="disabled")
        self.entry_value3.grid(row=3, column=1)

        tk.Label(master, text="Trabajo_Realizado:").grid(row=4, column=0, sticky="e")
        self.entry_value4 = tk.Entry(master, textvariable=self.value4, state="disabled")
        self.entry_value4.grid(row=4, column=1)

        self.btn_unlock = tk.Button(master, text="Buscar Equipo", command=self.unlock_fields)
        self.btn_unlock.grid(row=5, column=0, columnspan=2)

        self.btn_save = tk.Button(master, text="Guardar", command=self.save_to_excel)
        self.btn_save.grid(row=6, column=0, columnspan=2)

        self.btn_create_folder = tk.Button(master, text="Guardar imagen", command=self.create_folder)
        self.btn_create_folder.grid(row=7, column=0, columnspan=2)

    def unlock_fields(self):
        if self.reference.get():
            self.entry_value1.config(state="normal")
            self.entry_value2.config(state="normal")
            self.entry_value3.config(state="normal")
            self.entry_value4.config(state="normal")
        else:
            messagebox.showerror("Error", "Por favor, ingrese una referencia")

    def save_to_excel(self):
        if not all([self.reference.get(), self.value1.get(), self.value2.get(), self.value3.get(), self.value4.get()]):
            messagebox.showerror("Error", "Por favor, complete todos los campos")
            return

        # Verificar si el archivo existe, si no, crearlo
        if not os.path.exists(self.excel_file):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Datos"
            headers = ["EQUIPO", "SKU", "Nº RMA", "PRECIO", "TRABAJO REALIZADO"]
            ws.append(headers)
            self.format_headers(ws)
        else:
            wb = openpyxl.load_workbook(self.excel_file)
            ws = wb.active

        # Contar el número de filas con datos
        row_count = sum(1 for row in ws.iter_rows(min_row=2, max_col=1, values_only=True) if row[0])

        if row_count >= 36 and not self.reference_exists(ws):
            messagebox.showerror("Error", "Se ha alcanzado el límite de 36 entradas")
            return

        # Buscar si la referencia ya existe
        reference_exists = False
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0] == self.reference.get():
                reference_exists = True
                break

        if reference_exists:
            # Actualizar la fila existente
            for row in ws.iter_rows(min_row=2):
                if row[0].value == self.reference.get():
                    row[1].value = self.value1.get()
                    row[2].value = self.value2.get()
                    row[3].value = self.value3.get()
                    row[4].value = self.value4.get()
                    break
        else:
            # Añadir una nueva fila
            values = [self.reference.get(), self.value1.get(), self.value2.get(), self.value3.get(), self.value4.get()]
            ws.append(values)

        # Ajustar el ancho de las columnas
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        wb.save(self.excel_file)
        messagebox.showinfo("Éxito", "Datos guardados correctamente en el archivo Excel")
        
    def create_folder(self):
        if not self.reference.get():
            messagebox.showerror("Error", "Por favor, ingrese una referencia")
            return

        folder_name = self.reference.get()
        try:
            os.mkdir(folder_name)
            messagebox.showinfo("Éxito", f"Carpeta '{folder_name}' creada correctamente")
            
            # Abrir la carpeta automáticamente
            if os.name == 'nt':  # Para Windows
                os.startfile(folder_name)
            elif os.name == 'posix':  # Para macOS y Linux
                subprocess.call(('open', folder_name))
        except FileExistsError:
            messagebox.showerror("Error", f"La carpeta '{folder_name}' ya existe")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la carpeta: {str(e)}")

    def format_headers(self, worksheet):
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

    def reference_exists(self, worksheet):
        for row in worksheet.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0] == self.reference.get():
                return True
        return False
root = tk.Tk()
app = ExcelController(root)
root.mainloop()