import tkinter as tk
from tkinter import filedialog, messagebox
import hla_finder
import histomatch_analisis
import histomatch_reporte
import os
import shutil
import win32com.client
import pythoncom


class HistoMatchGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("HistoMatch")
        try:
            self.root.iconbitmap("mi_icono.ico")
        except:
            pass
        self.root.geometry("600x460")  # Reduje el alto para que se vea más compacto
        self.root.resizable(False, False)
        self.paciente_files = []
        self.donantes_files = []
        self.setup_ui()

    def setup_ui(self):
        # Header
        header = tk.Frame(self.root, bg="#2c3e50", height=60)
        header.pack(fill="x")
        tk.Label(header, text="HistoMatch - Prueba de Histocompatibilidad", fg="white",
                 bg="#2c3e50", font=("Arial", 16, "bold")).pack(pady=15)

        # Contenedor principal
        main_frame = tk.Frame(self.root, padx=30, pady=15)
        main_frame.pack(fill="both", expand=True)

        # SECCIÓN PACIENTE
        tk.Label(main_frame, text="1. Grupo Paciente:", font=("Arial", 10, "bold")).pack(anchor="w")
        btn_p = tk.Button(main_frame, text="Seleccionar archivos (PDF / Word)",
                          command=self.seleccionar_paciente, width=45)
        btn_p.pack(pady=5)
        self.lbl_p = tk.Label(main_frame, text="Ningún archivo seleccionado \n ", fg="gray", font=("Arial", 9))
        self.lbl_p.pack(anchor="w", padx=10)

        # SECCIÓN DONANTES
        tk.Label(main_frame, text="2. Grupo Donantes:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(10, 0))
        btn_d = tk.Button(main_frame, text="Seleccionar archivos (PDF / Word)",
                          command=self.seleccionar_donantes, width=45)
        btn_d.pack(pady=5)
        self.lbl_d = tk.Label(main_frame, text="Ningún archivo seleccionado", fg="gray", font=("Arial", 9))
        self.lbl_d.pack(anchor="w", padx=10)

        # SECCIÓN EJECUCIÓN (Espacios reducidos aquí)
        tk.Frame(main_frame, height=2, bd=1, relief="sunken").pack(fill="x", pady=15)

        self.btn_run = tk.Button(main_frame, text="INICIAR PROCESAMIENTO",
                                 bg="#27ae60", fg="white", font=("Arial", 11, "bold"),
                                 height=2, width=30, command=self.ejecutar_proceso)
        self.btn_run.pack(pady=(0, 5))  # Muy poco espacio abajo del botón

        # CRÉDITOS (Pegados al botón)
        lbl_creditos = tk.Label(main_frame,
                                text="Creado por Rodrigo Puertas y Valery Velasquez",
                                fg="#7f8c8d",
                                font=("Arial", 10, "bold"))
        lbl_creditos.pack(pady=30)

        # BARRA DE ESTADO
        self.status = tk.Label(self.root, text="Listo", bd=1, relief="sunken", anchor="w")
        self.status.pack(side="bottom", fill="x")

    def seleccionar_paciente(self):
        files = filedialog.askopenfilenames(title="Pacientes",
                                            filetypes=[("Documentos", ("*.pdf", "*.docx", "*.doc"))])
        if files:
            self.paciente_files = list(files)
            self.lbl_p.config(text=f"✓ {len(files)} cargado(s)", fg="green")

    def seleccionar_donantes(self):
        files = filedialog.askopenfilenames(title="Donantes",
                                            filetypes=[("Documentos", ("*.pdf", "*.docx", "*.doc"))])
        if files:
            self.donantes_files = list(files)
            self.lbl_d.config(text=f"✓ {len(files)} cargado(s)", fg="green")

    def ejecutar_proceso(self):
        if not self.paciente_files or not self.donantes_files:
            messagebox.showwarning("Aviso", "Selecciona los archivos primero.")
            return

        self.status.config(text="Iniciando Word para conversión...")
        self.root.update()

        word_app = None
        todos_preparados = []

        try:
            pythoncom.CoInitialize()
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0

            archivos_a_procesar = self.paciente_files + self.donantes_files

            for f in archivos_a_procesar:
                if f.lower().endswith((".doc", ".docx")):
                    self.status.config(text=f"Convirtiendo: {os.path.basename(f)}...")
                    self.root.update()

                    ruta_abs = os.path.abspath(f)
                    ruta_pdf = os.path.splitext(ruta_abs)[0] + ".pdf"

                    try:
                        doc = word_app.Documents.Open(ruta_abs, ReadOnly=True)
                        doc.SaveAs(ruta_pdf, FileFormat=17)
                        doc.Close(False)
                        todos_preparados.append(ruta_pdf)
                    except:
                        continue
                else:
                    todos_preparados.append(f)

            try:
                word_app.Quit()
                word_app = None
            except:
                pass

            self.status.config(text="Extrayendo datos HLA...")
            self.root.update()

            nombre_paciente = hla_finder.ejecutar_procesamiento(todos_preparados)

            if nombre_paciente:
                nombre_limpio = "".join([c for c in nombre_paciente if c.isalnum() or c in (' ', '_')]).strip()
                nombre_carpeta = f"HistoMatch_{nombre_limpio.replace(' ', '_')}"
                if not os.path.exists(nombre_carpeta): os.makedirs(nombre_carpeta)

                if os.path.exists("resultados_hla_final_limpio.csv"):
                    shutil.move("resultados_hla_final_limpio.csv",
                                os.path.join(nombre_carpeta, "resultados_hla_final_limpio.csv"))

                if histomatch_analisis.generar_informe(nombre_carpeta):
                    histomatch_reporte.generar_reporte(nombre_carpeta)
                    messagebox.showinfo("Éxito", f"Informe generado en: {nombre_carpeta}")
                    os.startfile(nombre_carpeta)
            else:
                messagebox.showerror("Error", "No se detectó el nombre del paciente.")

        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")
        finally:
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            self.status.config(text="Listo")


if __name__ == "__main__":
    root = tk.Tk()
    app = HistoMatchGUI(root)
    root.mainloop()
