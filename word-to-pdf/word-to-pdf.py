import tkinter as tk
from tkinter import ttk, filedialog
import os
import threading
import random
import time
from docx2pdf import convert


class WordToPdfConverter:
    DEFAULT_OUT = r"C:\Users\28570\Desktop\PDF转换结果"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Word转PDF小工具")
        self.root.geometry("580x220")
        self.root.resizable(False, False)

        self.word_path = ""
        self.pdf_path  = ""

        os.makedirs(self.DEFAULT_OUT, exist_ok=True)
        self.build_ui()

    # -------------------- 界面 --------------------
    def build_ui(self):
        # 标题
        tk.Label(self.root, text="Word转PDF小工具", font=("微软雅黑", 14)).pack(pady=8)

        # Word 文件（一行）
        row1 = tk.Frame(self.root)
        row1.pack(fill="x", padx=20, pady=4)
        tk.Label(row1, text="Word文件：", width=10, anchor="w").pack(side="left")
        self.word_entry = tk.Entry(row1, width=45)
        self.word_entry.pack(side="left", padx=(0, 5))
        tk.Button(row1, text="浏览...", width=8, command=self.select_word).pack(side="left")

        # 输出目录（一行）
        row2 = tk.Frame(self.root)
        row2.pack(fill="x", padx=20, pady=4)
        tk.Label(row2, text="输出目录：", width=10, anchor="w").pack(side="left")
        self.out_entry = tk.Entry(row2, width=45)
        self.out_entry.pack(side="left", padx=(0, 5))
        self.out_entry.insert(0, self.DEFAULT_OUT)
        tk.Button(row2, text="浏览...", width=8, command=self.select_out).pack(side="left")

        # 进度条
        self.progress = ttk.Progressbar(self.root, mode="determinate", length=540)
        self.progress.pack(pady=8)

        # 按钮行（开始 | 打开）
        row3 = tk.Frame(self.root)
        row3.pack(fill="x", padx=20, pady=6)
        self.convert_btn = tk.Button(row3,
                                     text="开始转换",
                                     command=self.start_convert,
                                     bg="#4CAF50", fg="white", width=15)
        self.convert_btn.pack(side="left")
        self.open_btn = tk.Button(row3,
                                  text="打开文件",
                                  command=self.open_pdf,
                                  state="disabled", width=15)
        self.open_btn.pack(side="right")

    # -------------------- 选择 --------------------
    def select_word(self):
        file = filedialog.askopenfilename(filetypes=[("Word 文件", "*.docx *.doc")])
        if file:
            self.word_path = file
            self.word_entry.delete(0, tk.END)
            self.word_entry.insert(0, file)

    def select_out(self):
        folder = filedialog.askdirectory()
        if folder:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, folder)

    # -------------------- 转换 --------------------
    def start_convert(self):
        if not self.word_path:
            return
        self.convert_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self.do_work, daemon=True).start()

    def do_work(self):
        try:
            out_dir = self.out_entry.get()
            os.makedirs(out_dir, exist_ok=True)
            base = os.path.splitext(os.path.basename(self.word_path))[0]
            self.pdf_path = os.path.join(out_dir, f"{base}.pdf")

            # 后台转换
            convert(self.word_path, self.pdf_path)

            # 灵活进度条：随机步长+短延时，模拟“柑橘不同步”
            val = 0
            while val < 100:
                delta = random.randint(3, 10)
                val = min(val + delta, 100)
                self.progress["value"] = val
                time.sleep(random.uniform(0.02, 0.08))
            self.open_btn.config(state="normal")
        except Exception:
            pass
        finally:
            self.convert_btn.config(state="normal")

    # -------------------- 打开结果 --------------------
    def open_pdf(self):
        if os.path.isfile(self.pdf_path):
            os.startfile(self.pdf_path)


# -------------------- 启动 --------------------
if __name__ == "__main__":
    root = tk.Tk()
    WordToPdfConverter(root)
    root.mainloop()