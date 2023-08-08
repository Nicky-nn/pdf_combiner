import PyPDF2
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from threading import Thread
import subprocess

class PDFCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Combinar PDFs")

        self.selected_pdf_files = []

        self.combine_button = tk.Button(self.root, text="Combinar PDFs", command=self.combine_pdfs)
        self.combine_button.pack(pady=10)

        self.progress_label = tk.Label(self.root, text="Progreso:")
        self.progress_label.pack()

        self.progress_bar = ttk.Progressbar(self.root, orient='horizontal', length=300, mode='determinate')
        self.progress_bar.pack(pady=5)

        self.cancel_button = tk.Button(self.root, text="Cancelar", command=self.cancel_combination, state=tk.DISABLED)
        self.cancel_button.pack(pady=5)

        self.result_label = tk.Label(self.root, text="", fg="green")
        self.result_label.pack()

        self.footer_label = tk.Label(self.root, text="© 2023 Integrate IT Solutions | Desarrollado por Nick Russell", fg="gray")
        self.footer_label.pack(pady=10)

        self.cancelled = False

    def combine_pdfs(self):
        self.selected_pdf_files = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])

        if not self.selected_pdf_files:
            return

        output_filename = 'pdf_combinado.pdf'

        self.progress_bar['maximum'] = len(self.selected_pdf_files)
        self.progress_bar['value'] = 0

        self.progress_label.config(text="Progreso: 0%")
        self.result_label.config(text="", fg="black")

        self.combine_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)

        # Iniciar el proceso de combinación en un hilo separado
        thread = Thread(target=self.combine_pdf_files, args=(output_filename,))
        thread.start()

    def combine_pdf_files(self, output_filename):
        pdf_merger = PyPDF2.PdfMerger()

        for index, pdf_file in enumerate(self.selected_pdf_files, start=1):
            if self.cancelled:
                self.cleanup_after_cancel()
                return

            pdf_merger.append(pdf_file)
            self.progress_bar.step(1)
            progress_percent = (index / len(self.selected_pdf_files)) * 100
            self.progress_label.config(text=f"Progreso: {progress_percent:.2f}%")
            self.root.update_idletasks()

        with open(output_filename, 'wb') as output_pdf:
            pdf_merger.write(output_pdf)

        self.result_label.config(text=f'Se han combinado {len(self.selected_pdf_files)} archivos PDF en "{output_filename}"', fg="green")
        self.cleanup_after_completion(output_filename)

    def cancel_combination(self):
        self.cancelled = True
        self.combine_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)

    def cleanup_after_cancel(self):
        self.cancelled = False
        self.progress_label.config(text="Progreso:")
        self.progress_bar['value'] = 0
        self.result_label.config(text="Combinación cancelada", fg="red")

    def cleanup_after_completion(self, output_filename):
        self.progress_label.config(text="Progreso:")
        self.progress_bar['value'] = 0
        self.combine_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)

        # Obtener la ubicación del archivo resultante
        output_path = os.path.abspath(output_filename)
        output_folder = os.path.dirname(output_path)

        # Abrir la carpeta en el explorador de archivos
        subprocess.Popen(['explorer', output_folder])

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCombinerApp(root)
    root.mainloop()
