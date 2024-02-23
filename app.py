import tkinter as tk
from tkinter import ttk, filedialog
from classes.funciones import generate_alldata_joints_fromIFC, transform_boxes_info_for_bom_excel_and_generate
from threading import Thread
import os

class BOMGeneratorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Generador de BOM de Joints")
        self.master.geometry("400x350")
        self.master.configure(bg="#497767")

        self.label = ttk.Label(master, text="Generador de BOM de Joints", font=("Helvetica", 16), background="#497767", foreground="white")
        self.label.pack(pady=20)

        self.file_path = ""
        self.progress_label = ttk.Label(master, text="", background="#497767", foreground="white")
        self.progress_label.pack(pady=10)

        self.select_button = ttk.Button(master, text="Selecciona archivo IFC", command=self.show_file_dialog)
        self.select_button.pack(pady=10)

        self.start_button = ttk.Button(master, text="Comenzar", command=self.start_processing, state=tk.DISABLED)
        self.start_button.pack(pady=10)

        self.open_excel_button = ttk.Button(master, text="Abrir Excel", command=self.open_excel, state=tk.DISABLED)
        self.open_excel_button.pack(pady=10)

    def show_file_dialog(self):
        self.file_path = filedialog.askopenfilename(title="Seleccionar archivo IFC", filetypes=[("Archivos IFC", "*.ifc")])
        if self.file_path:
            self.progress_label.config(text=f"Archivo seleccionado: {self.file_path}")
            self.start_button.config(state=tk.NORMAL)

    def start_processing(self):
        self.start_button.config(state=tk.DISABLED)
        self.progress_label.config(text="Procesando, por favor espera...")

        # Ejecutar las funciones en un hilo para no bloquear la interfaz gráfica
        process_thread = Thread(target=self.generate_bom)
        process_thread.start()

    def generate_bom(self):
        try:
            boxes_objects, joints_bom_lines, inferredconnections_bom_lines, modeledconnections_bom_lines = generate_alldata_joints_fromIFC(self.file_path)
            transform_boxes_info_for_bom_excel_and_generate(boxes_objects, joints_bom_lines, inferredconnections_bom_lines, modeledconnections_bom_lines, self.file_path)

            self.progress_label.config(text="BOM generado exitosamente.")
            self.open_excel_button.config(state=tk.NORMAL)
        except Exception as e:
            self.progress_label.config(text=f"Error: {str(e)}")

        self.start_button.config(state=tk.NORMAL)

    def open_excel(self):
        excel_path = os.path.splitext(self.file_path)[0] + ".xlsx"
        try:
            os.startfile(excel_path)
        except Exception as e:
            self.progress_label.config(text=f"Error al abrir el archivo Excel: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = BOMGeneratorApp(root)
    root.mainloop()

