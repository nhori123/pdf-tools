import os
import shutil
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image, ImageTk
import pdf2image
import sys

class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Splitter")
        self.root.geometry("600x400")
        
        # 実行可能ファイルからのパスを考慮してアイコンを設定
        icon_path = os.path.join(sys._MEIPASS, "pdf-splitter.ico") if hasattr(sys, "_MEIPASS") else "pdf-splitter.ico"
        self.root.iconbitmap(icon_path)

        self.filename = ""
        self.pdf_reader = None
        self.page_images = []
        self.rotated_pages = {}
        self.selected_page = None
        self.labels = []

        self.create_widgets()
        
    def create_widgets(self):
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        self.open_button = ttk.Button(button_frame, text="Open PDF", command=self.open_pdf)
        self.open_button.pack(side=tk.LEFT, padx=5)
        
        self.rotate_button = ttk.Button(button_frame, text="Rotate", command=self.rotate_page, state=tk.DISABLED)
        self.rotate_button.pack(side=tk.LEFT, padx=5)

        self.save_button = ttk.Button(button_frame, text="Save Split PDFs", command=self.save_split_pdfs, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
    def open_pdf(self):
        self.filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.filename:
            self.pdf_reader = PdfReader(self.filename)
            self.create_temp_dir()
            self.show_preview()
            self.save_button.config(state=tk.NORMAL)
            self.rotate_button.config(state=tk.NORMAL)
        
    def create_temp_dir(self):
        pdf_dir = os.path.dirname(self.filename)
        pdf_basename = os.path.basename(self.filename)
        pdf_name, pdf_ext = os.path.splitext(pdf_basename)
        
        self.temp_dir = os.path.join(pdf_dir, f".{pdf_name}_temp")
        os.makedirs(self.temp_dir, exist_ok=True)
        
    def show_preview(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.labels = []
        self.page_images = pdf2image.convert_from_path(self.filename, output_folder=self.temp_dir, first_page=1, last_page=len(self.pdf_reader.pages))
        
        for page_num, img in enumerate(self.page_images):
            img.thumbnail((400, 400))
            img_tk = ImageTk.PhotoImage(img)
            label = tk.Label(self.scrollable_frame, image=img_tk, borderwidth=0, relief="solid")
            label.image = img_tk
            label.pack(side=tk.TOP, pady=5)
            label.bind("<Button-1>", lambda e, num=page_num: self.select_page(num))
            self.labels.append(label)
        
    def select_page(self, page_num):
        self.selected_page = page_num
        for label in self.labels:
            label.config(borderwidth=0)  # Reset all labels border
        self.labels[page_num].config(borderwidth=2)  # Highlight selected label
        
    def rotate_page(self):
        if self.selected_page is not None:
            page_num = self.selected_page
            self.page_images[page_num] = self.page_images[page_num].rotate(-90, expand=True)  # Rotate clockwise
            self.rotated_pages[page_num] = self.rotated_pages.get(page_num, 0) + 90
            self.update_preview()
        
    def update_preview(self):
        for page_num, img in enumerate(self.page_images):
            img.thumbnail((400, 400))
            img_tk = ImageTk.PhotoImage(img)
            self.labels[page_num].config(image=img_tk)
            self.labels[page_num].image = img_tk
            if page_num == self.selected_page:
                self.labels[page_num].config(borderwidth=2)  # Maintain highlight for rotated page
        
    def save_split_pdfs(self):
        pdf_dir = os.path.dirname(self.filename)
        pdf_basename = os.path.basename(self.filename)
        pdf_name, pdf_ext = os.path.splitext(pdf_basename)
        
        output_dir = os.path.join(pdf_dir, pdf_name)
        os.makedirs(output_dir, exist_ok=True)
        
        for page_num in range(len(self.pdf_reader.pages)):
            pdf_writer = PdfWriter()
            page = self.pdf_reader.pages[page_num]
            
            rotation = self.rotated_pages.get(page_num, 0)
            if rotation != 0:
                page = page.rotate(rotation)  # Rotate clockwise
            
            pdf_writer.add_page(page)
            
            output_filename = os.path.join(output_dir, f"{pdf_name}_p{page_num+1}{pdf_ext}")
            with open(output_filename, 'wb') as output_file:
                pdf_writer.write(output_file)
        
        tk.messagebox.showinfo("Success", "PDF files have been split and saved.")
        self.cleanup_temp_dir()
        
    def cleanup_temp_dir(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)
        
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFSplitterApp(root)
    root.mainloop()
