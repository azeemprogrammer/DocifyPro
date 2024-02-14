import PyPDF2
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from docx import Document
import threading
import os

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DocifyPro - PDF to Word Converter")

        self.file_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()

        # Frame for inputs
        input_frame = tk.Frame(self.root)
        input_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")

        # Software name label
        tk.Label(input_frame, text="DocifyPro", font=("Arial", 16)).grid(row=0, columnspan=2, pady=10)

        # PDF File input
        tk.Label(input_frame, text="PDF File:").grid(row=1, column=0, pady=5, sticky="w")
        self.file_entry = tk.Entry(input_frame, textvariable=self.file_path, width=50)
        self.file_entry.grid(row=1, column=1, pady=5, padx=5, sticky="we")
        self.file_entry.insert(0, "Select PDF file")

        # Output Folder input
        tk.Label(input_frame, text="Output Folder:").grid(row=2, column=0, pady=5, sticky="w")
        self.output_entry = tk.Entry(input_frame, textvariable=self.output_folder_path, width=50)
        self.output_entry.grid(row=2, column=1, pady=5, padx=5, sticky="we")
        self.output_entry.insert(0, "Select output folder")

        # Browse buttons
        tk.Button(input_frame, text="Browse", command=self.browse_file).grid(row=1, column=2, padx=5, pady=5, sticky="we")
        tk.Button(input_frame, text="Browse", command=self.browse_output_folder).grid(row=2, column=2, padx=5, pady=5, sticky="we")

        # Convert button
        tk.Button(self.root, text="Convert", command=self.convert).grid(row=1, column=0, padx=10, pady=10)

        # Progress bar
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.grid(row=2, padx=10, pady=5)

        # Labels for progress info
        self.total_pages_label = tk.Label(self.root, text="")
        self.total_pages_label.grid(row=3, padx=10, pady=5, sticky="w")
        self.converted_pages_label = tk.Label(self.root, text="")
        self.converted_pages_label.grid(row=4, padx=10, pady=5, sticky="w")
        self.remaining_pages_label = tk.Label(self.root, text="")
        self.remaining_pages_label.grid(row=5, padx=10, pady=5, sticky="w")

        # Set window size and make it fixed
        self.root.geometry("500x300")
        self.root.resizable(False, False)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.file_path.set(file_path)

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_path.set(folder_path)

    def update_progress(self, current, total):
        self.progress_bar["value"] = (current / total) * 100
        self.root.update_idletasks()

    def convert(self):
        file_path = self.file_path.get()
        output_folder = self.output_folder_path.get()

        if not file_path or not output_folder:
            messagebox.showerror("Error", "Please select PDF file and output folder.")
            return

        try:
            pdf_reader = PyPDF2.PdfReader(file_path)
            total_pages = len(pdf_reader.pages)

            self.total_pages_label.config(text=f"Total Pages: {total_pages}")
            self.converted_pages_label.config(text="Converted Pages: 0")
            self.remaining_pages_label.config(text="Remaining Pages: 0")

            def conversion_worker():
                for page_num in range(total_pages):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()

                    doc = Document()
                    doc.add_paragraph(text)

                    output_path = os.path.join(output_folder, f"page_{page_num + 1}.docx")
                    doc.save(output_path)

                    self.converted_pages_label.config(text=f"Converted Pages: {page_num + 1}")
                    self.remaining_pages_label.config(text=f"Remaining Pages: {total_pages - page_num - 1}")
                    self.update_progress(page_num + 1, total_pages)

            threading.Thread(target=conversion_worker).start()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

def main():
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
