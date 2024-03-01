import tkinter as tk
from tkinter import ttk, filedialog
from classes.funciones import generate_alldata_joints_fromIFC, transform_boxes_info_for_bom_excel_and_generate
from threading import Thread
import os
import io
import sys
from PIL import Image, ImageTk
from pathlib import Path

class BOMGeneratorApp:
    def __init__(self, master):
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

        self.master = master
        self.master.title("Cost & BOM Joints Generator v1.00")
        self.master.geometry("550x550")
        self.master.configure(bg="#101E15")

        self.font_style = ("Ingram Mono", 16)

        self.label = ttk.Label(master, text="Cost & BOM Joints Generator", font=self.font_style, background="#101E15", foreground="#7FEA4E")
        self.label.pack(pady=10)

        if getattr(sys, 'frozen', False):
            base_path = Path(sys._MEIPASS)
        else:
            base_path = Path(os.path.abspath("."))

        image_path = base_path / "logoverde.png"

        if image_path.is_file():
            self.logo_image = ImageTk.PhotoImage(Image.open(str(image_path)))
            self.logo_label = tk.Label(master, image=self.logo_image, background="#101E15")
            self.logo_label.pack(pady=10)
        else:
            print("Error: No se encontró la imagen.")

        icon_path = base_path / "icono.ico"

        if icon_path.is_file():
            self.master.iconbitmap(str(icon_path))
        else:
            print("Error: No se encontró el icono.")

        self.file_path = ""
        self.progress_text_var = tk.StringVar()
        self.progress_text_var.set("Select an IFC file and click START")
        self.progress_label = ttk.Label(master, textvariable=self.progress_text_var, font=("Ingram Mono", 8), background="#101E15", foreground="#7FEA4E")
        self.progress_label.pack(pady=10)

        button_font_style = ("Ingram Mono", 8)

        self.select_button = tk.Button(master, text="Select IFC file", command=self.show_file_dialog, font=button_font_style, bg="#3A6739", fg="#EEE8D3", padx=10, pady=5)
        self.select_button.pack(pady=10)

        self.start_button = tk.Button(master, text="Start", command=self.start_processing, state=tk.DISABLED, font=button_font_style, bg="#3A6739", fg="#EEE8D3", padx=10, pady=5)
        self.start_button.pack(pady=10)

        self.open_excel_button = tk.Button(master, text="Open Excel file", command=self.open_excel, state=tk.DISABLED, font=button_font_style, bg="#3A6739", fg="#EEE8D3", padx=10, pady=5)
        self.open_excel_button.pack(pady=10)

        self.readme_button = tk.Button(master, text="Readme", command=self.open_readme, font=button_font_style, bg="#3A6739", fg="#EEE8D3", padx=10, pady=5)
        self.readme_button.pack(pady=10)

        self.stdout_buffer = io.StringIO()
        sys.stdout = self.stdout_buffer

        self.terminal_text = tk.Text(master, wrap="word", font=("Ingram Mono", 8), height=10, state=tk.DISABLED, background="#101E15", foreground="#7FEA4E")
        self.terminal_text.pack(pady=10, padx=20, fill="both", expand=True)

    def show_file_dialog(self):
        self.file_path = filedialog.askopenfilename(title="Select IFC file", filetypes=[("IFC Files", "*.ifc")])
        filedialog_font_style = ("Ingram Mono", 8)
        if self.file_path:
            self.progress_text_var.set(f"File selected: {self.file_path}")
            self.start_button.config(state=tk.NORMAL)
            self.select_button.config(state=tk.DISABLED)

    def start_processing(self):
        self.start_button.config(state=tk.DISABLED)
        self.progress_text_var.set("Processing, please wait...")

        process_thread = Thread(target=self.generate_bom)
        process_thread.start()

        self.master.after(100, self.update_terminal)

    def generate_bom(self):
        try:
            boxes_objects, joints_bom_lines, inferredconnections_bom_lines, modeledconnections_bom_lines = generate_alldata_joints_fromIFC(self.file_path)

            self.progress_text_var.set("Processing, please wait...")

            transform_boxes_info_for_bom_excel_and_generate(boxes_objects, joints_bom_lines, inferredconnections_bom_lines, modeledconnections_bom_lines, self.file_path)

            self.progress_text_var.set("BOM generated successfully  \(❛ ͜ʖ❛)/")
            self.open_excel_button.config(state=tk.NORMAL)
            self.readme_button.config(state=tk.NORMAL)
        except Exception as e:
            self.progress_text_var.set(f"Error: {str(e)}")

        self.start_button.config(state=tk.DISABLED)

    def update_terminal(self):
        output_text = self.stdout_buffer.getvalue()

        self.terminal_text.config(state=tk.NORMAL)
        self.terminal_text.delete(1.0, tk.END)
        self.terminal_text.insert(tk.END, output_text)
        self.terminal_text.config(state=tk.DISABLED)

        self.master.after(100, self.update_terminal)

    def show_in_terminal(self, line):
        self.terminal_text.config(state=tk.NORMAL)
        self.terminal_text.insert(tk.END, line + "\n")
        self.terminal_text.config(state=tk.DISABLED)

    def open_excel(self):
        excel_path = os.path.splitext(self.file_path)[0] + ".xlsx"
        try:
            os.startfile(excel_path)
        except Exception as e:
            self.progress_text_var.set(f"Error opening Excel file: {str(e)}")

    def open_readme(self):
        readme_path = "leeme.txt"
        try:
            with open(readme_path, "r", encoding="utf-8") as readme_file:
                readme_content = readme_file.read()
                self.show_readme_dialog(readme_content)
        except Exception as e:
            self.progress_text_var.set(f"Error opening Readme file: {str(e)}")

    def show_readme_dialog(self, content):
        readme_dialog = tk.Toplevel(self.master)
        readme_dialog.title("Readme")
        readme_dialog.geometry("700x700")

        readme_text = tk.Text(readme_dialog, wrap="word", font=("Ingram Mono", 8))
        readme_text.insert(tk.END, content)
        readme_text.pack(expand=True, fill="both", padx=10, pady=10)

        close_button = tk.Button(readme_dialog, text="Close", command=readme_dialog.destroy, font=("Ingram Mono", 8), bg="#3A6739", fg="#EEE8D3", padx=10, pady=5)
        close_button.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = BOMGeneratorApp(root)
    root.mainloop()