# このコードでは、ファイルリストエリア、プレビューエリア、操作ボタン、出力ファイルエリアによって構成されています。
# ユーザーはPDFファイルを追加・削除し、結合するファイルの保存先を指定して、結合ボタンをクリックするだけで、
# 複数のPDFファイルを1つに結合できます。
# PDFの回転にも対応しています。ただし、複数ページあるPDFファイルの場合、すべてのPDFファイルが回転します。
# 既知の不具合として、作業用の一時フォルダ削除されずにのこってしまいます。
# 一時フォルダはCドライブ直下に「.pdf-marge」から始まるフォルダ名で保存されるため、ツール利用後は手動で削除を行ってください。

import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
import os
import sys
from PIL import Image, ImageTk
import fitz  # PyMuPDF
import datetime
import shutil
import atexit

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF結合ツール")
        self.root.geometry("1000x700")
        self.root.iconbitmap(self.resource_path("pdf-merge.ico"))

        self.file_paths = []
        self.temp_dirs = []

        # フレームの作成
        frame = tk.Frame(root)
        frame.pack(pady=10)

        # ファイルリストを表示するリストボックスとプレビュー用のキャンバス
        self.file_list = tk.Listbox(frame, width=50, height=20)
        self.file_list.grid(row=0, column=0, padx=10)

        self.preview_canvas = tk.Canvas(frame, width=500, height=500)
        self.preview_canvas.grid(row=0, column=1, padx=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        add_btn = tk.Button(btn_frame, text="PDFファイルを追加", command=self.add_file)
        add_btn.grid(row=0, column=0, padx=5)

        remove_btn = tk.Button(btn_frame, text="選択されたファイルを削除", command=self.remove_file)
        remove_btn.grid(row=0, column=1, padx=5)

        up_btn = tk.Button(btn_frame, text="上へ", command=self.move_up)
        up_btn.grid(row=0, column=2, padx=5)

        down_btn = tk.Button(btn_frame, text="下へ", command=self.move_down)
        down_btn.grid(row=0, column=3, padx=5)

        rotate_btn = tk.Button(btn_frame, text="回転", command=self.rotate_pdf)
        rotate_btn.grid(row=0, column=4, padx=5)

        path_frame = tk.Frame(root)
        path_frame.pack(pady=10)
        self.output_entry = tk.Entry(path_frame, width=50)
        self.output_entry.grid(row=0, column=0)
        self.output_entry.insert(0, "結合後のPDFファイルの保存先を入力...")

        browse_btn = tk.Button(path_frame, text="参照", command=self.browse_output)
        browse_btn.grid(row=0, column=1, padx=5)

        merge_btn = tk.Button(root, text="結合", command=self.merge_pdfs)
        merge_btn.pack(pady=5)

        self.file_list.bind("<<ListboxSelect>>", self.show_preview)

        atexit.register(self.cleanup_temp_dirs)

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def add_file(self):
        files = filedialog.askopenfilenames(filetypes=[("PDFファイル", "*.pdf")])
        for file in files:
            self.file_paths.append(file)
            self.file_list.insert(tk.END, os.path.basename(file))
            self.show_preview()

    def remove_file(self):
        selected = self.file_list.curselection()
        for i in reversed(selected):
            self.file_list.delete(i)
            del self.file_paths[i]
            self.preview_canvas.delete("all")

    def move_up(self):
        selected = self.file_list.curselection()
        if not selected:
            return
        for i in selected:
            if i == 0:
                continue
            file_name = self.file_list.get(i)
            self.file_list.delete(i)
            self.file_list.insert(i-1, file_name)

            file_path = self.file_paths.pop(i)
            self.file_paths.insert(i-1, file_path)

        self.file_list.selection_set(selected[0]-1)
        self.show_preview()

    def move_down(self):
        selected = self.file_list.curselection()
        if not selected:
            return
        for i in reversed(selected):
            if i == len(self.file_list.get(0, tk.END)) - 1:
                continue
            file_name = self.file_list.get(i)
            self.file_list.delete(i)
            self.file_list.insert(i+1, file_name)

            file_path = self.file_paths.pop(i)
            self.file_paths.insert(i+1, file_path)

        self.file_list.selection_set(selected[0]+1)
        self.show_preview()

    def browse_output(self):
        file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDFファイル", "*.pdf")])
        if file:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file)

    def show_preview(self, event=None):
        selected = self.file_list.curselection()
        if not selected:
            return
        file_path = self.file_paths[selected[0]]

        # PDFファイルの1ページ目を画像に変換してキャンバスに表示
        doc = fitz.open(file_path)
        page = doc.load_page(0)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((500, 500))
        self.img_tk = ImageTk.PhotoImage(img)
        self.preview_canvas.create_image(0, 0, anchor="nw", image=self.img_tk)

    def rotate_pdf(self):
        selected = self.file_list.curselection()
        if not selected:
            messagebox.showwarning("警告", "PDFファイルを選択してください")
            return

        file_path = self.file_paths[selected[0]]

        # 一時フォルダの作成
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        temp_dir = f"C:/.pdf-merge{timestamp}"
        os.makedirs(temp_dir, exist_ok=True)
        self.temp_dirs.append(temp_dir)

        temp_file_path = os.path.join(temp_dir, os.path.basename(file_path))

        with open(file_path, "rb") as infile:
            reader = PyPDF2.PdfReader(infile)
            writer = PyPDF2.PdfWriter()

            for page in reader.pages:
                page.rotate(90)
                writer.add_page(page)

            with open(temp_file_path, "wb") as outfile:
                writer.write(outfile)

        self.file_paths[selected[0]] = temp_file_path
        self.show_preview()

    def merge_pdfs(self):
        if not self.file_paths:
            messagebox.showwarning("警告", "PDFファイルを追加してください")
            return

        output_path = self.output_entry.get()
        if not output_path or output_path == "結合後のPDFファイルの保存先を入力...":
            messagebox.showwarning("警告", "保存先を指定してください")
            return

        pdf_merger = PyPDF2.PdfMerger()
        for file in self.file_paths:
            pdf_merger.append(file)

        with open(output_path, 'wb') as f:
            pdf_merger.write(f)

        messagebox.showinfo("情報", "PDFの結合が完了しました！")

    def cleanup_temp_dirs(self):
        for temp_dir in self.temp_dirs:
            shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
