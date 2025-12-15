# -*- coding: utf-8 -*-
"""
office2pdf.py  ——  Office→PDF 小工具（MS/WPS/LibreOffice 自动识别）
Python 3.8+  |  pip install pywin32
"""
import os
import sys
import subprocess
import threading
import time
import winreg
from pathlib import Path
from urllib.request import urlretrieve
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from queue import Queue, Empty
import pythoncom  # 子线程 STA

# -------------------- 常量 --------------------
LO_URL = (
    "https://downloadarchive.documentfoundation.org/"
    "libreoffice/old/7.6.4.1/win/x86_64/LibreOffice_7.6.4.1_Win_x86_64.msi"
)
LO_FILTERS = {
    ".docx": "writer_pdf_Export",
    ".doc": "writer_pdf_Export",
    ".xlsx": "calc_pdf_Export",
    ".xls": "calc_pdf_Export",
    ".pptx": "impress_pdf_Export",
    ".ppt": "impress_pdf_Export",
}

# -------------------- 检测本机 Office --------------------
def which_office():
    try:
        import win32com.client
        win32com.client.Dispatch("Word.Application")
        return "MS"
    except Exception:
        pass
    for progid in ("kwps.Application", "wps.Application"):
        try:
            import win32com.client
            win32com.client.Dispatch(progid)
            return "WPS"
        except Exception:
            continue
    return "LO" if is_lo_available() else "None"

def is_lo_available():
    try:
        subprocess.check_output(["soffice", "--version"], stderr=subprocess.DEVNULL)
        return True
    except Exception:
        if sys.platform == "win32":
            return find_lo_from_reg_or_disk() is not None
        return False

def find_lo_from_reg_or_disk():
    for key_path in [
        r"SOFTWARE\LibreOffice\UNO\InstallPath",
        r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath",
    ]:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                candidate = os.path.join(winreg.QueryValueEx(key, "")[0], "soffice.exe")
                if os.path.isfile(candidate):
                    return candidate
        except Exception:
            continue
    for candidate in [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]:
        if os.path.isfile(candidate):
            return candidate
    return None

# -------------------- 真正干活的转换函数（纯后台） --------------------
def _do_convert(file_path, out_dir, bridge):
    """
    无论 MS / WPS / LO，都在本函数里跑，跑完后通过 bridge 回传进度。
    返回值: (pdf_path or None, error_str or None)
    """
    suffix = Path(file_path).suffix.lower()
    typ = {".docx": "word", ".doc": "word",
           ".xlsx": "excel", ".xls": "excel",
           ".pptx": "ppt", ".ppt": "ppt"}.get(suffix, "")
    if not typ:
        return None, "不支持的文件类型"

    engine = which_office()
    if engine not in ("MS", "WPS", "LO"):
        return None, "本机未检测到 MS/WPS/LibreOffice"

    pdf_path = os.path.join(out_dir, f"{Path(file_path).stem}.pdf")
    os.makedirs(out_dir, exist_ok=True)

    # 1. 启动阶段 0-30（最多 1.5 s）
    for i in range(30):
        try:
            if engine == "LO":
                bridge.update(30)
                break
            else:
                import win32com.client as win32
                progid = {
                    "MS": {"word": "Word.Application", "excel": "Excel.Application", "ppt": "PowerPoint.Application"},
                    "WPS": {"word": "kwps.Application", "excel": "ket.Application", "ppt": "kwpp.Application"}
                }[engine][typ]
                win32.Dispatch(progid)
                bridge.update(30)
                break
        except Exception:
            time.sleep(0.001)
            bridge.update(i)

    # 2. 转换阶段 30-100
    if engine == "LO":
        bridge.update(35)
        cmd = [find_lo_from_reg_or_disk() or "soffice",
               "--headless", "--convert-to", f"pdf:{LO_FILTERS[suffix]}",
               "--outdir", out_dir, os.path.abspath(file_path)]
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        for i in range(35, 96):
            if proc.poll() is not None:
                bridge.update(96)
                break
            time.sleep(0.001)
            bridge.update(i)
        if proc.returncode != 0:
            return None, "LibreOffice 转换失败"
    else:
        done = threading.Event()
        error = []

        def _bg():
            try:
                if engine == "MS":
                    convert_ms(file_path, out_dir, typ)
                else:
                    convert_wps(file_path, out_dir, typ)
            except Exception as e:
                error.append(str(e))
            done.set()

        threading.Thread(target=_bg, daemon=True).start()
        for i in range(35, 96):
            if done.is_set():
                bridge.update(96)
                break
            time.sleep(0.08)
            bridge.update(i)
        if error:
            return None, error[0]

    # 3. 收尾
    for i in range(96, 101):
        bridge.update(i)
        time.sleep(0.02)
    return pdf_path, None

# -------------------- MS / WPS / LO 核心转换 --------------------
def convert_ms(path, out_dir, typ):
    import win32com.client as win32
    app = win32.Dispatch(
        {"word": "Word.Application", "excel": "Excel.Application", "ppt": "PowerPoint.Application"}[typ]
    )
    src = os.path.abspath(path)
    pdf = os.path.join(out_dir, f"{Path(path).stem}.pdf")
    if typ == "word":
        doc = app.Documents.Open(src)
        doc.SaveAs2(pdf, FileFormat=17)
    elif typ == "excel":
        wb = app.Workbooks.Open(src)
        wb.ExportAsFixedFormat(0, pdf)
    elif typ == "ppt":
        pres = app.Presentations.Open(src)
        pres.SaveAs(pdf, 32)
    # 不 Close 不 Quit
    return pdf

def convert_wps(path, out_dir, typ):
    import win32com.client as win32
    if typ == "word":
        for p in ("kwps.Application", "wps.Application"):
            try:
                app = win32.Dispatch(p)
                break
            except Exception:
                continue
        else:
            raise RuntimeError("WPS Word 启动失败")
        doc = app.Documents.Open(os.path.abspath(path))
        pdf = os.path.join(out_dir, f"{Path(path).stem}.pdf")
        doc.ExportAsFixedFormat(pdf, 17)
        return pdf
    if typ == "excel":
        for p in ("ket.Application", "et.Application"):
            try:
                app = win32.Dispatch(p)
                break
            except Exception:
                continue
        else:
            raise RuntimeError("WPS 表格 启动失败")
        wb = app.Workbooks.Open(os.path.abspath(path))
        pdf = os.path.join(out_dir, f"{Path(path).stem}.pdf")
        wb.ExportAsFixedFormat(0, pdf)
        return pdf
    if typ == "ppt":
        for p in ("kwpp.Application", "wpp.Application"):
            try:
                app = win32.Dispatch(p)
                break
            except Exception:
                continue
        else:
            raise RuntimeError("WPS 演示 启动失败")
        pres = app.Presentations.Open(os.path.abspath(path))
        pdf = os.path.join(out_dir, f"{Path(path).stem}.pdf")
        pres.SaveAs(pdf, 32)
        return pdf

def convert_lo(path, out_dir):
    suffix = Path(path).suffix.lower()
    filter_name = LO_FILTERS.get(suffix)
    if not filter_name:
        raise RuntimeError("LibreOffice 不支持此文件类型")
    exe = find_lo_from_reg_or_disk() or "soffice"
    cmd = [exe, "--headless", "--convert-to", f"pdf:{filter_name}",
           "--outdir", out_dir, os.path.abspath(path)]
    subprocess.run(cmd, check=True, capture_output=True)
    pdf_path = os.path.join(out_dir, f"{Path(path).stem}.pdf")
    for _ in range(50):
        if os.path.isfile(pdf_path):
            return pdf_path
        time.sleep(0.1)
    raise FileNotFoundError("LibreOffice 未生成 PDF")

# -------------------- 自动下载 LibreOffice --------------------
def download_install_lo(parent):
    temp_msi = os.path.join(os.getenv("TEMP"), "LibreOffice_setup.msi")
    def reporthook(b, s, t):
        parent.progress["value"] = min(int(b * s * 100 / t), 100)
        parent.root.update()
    parent.progress.start()
    try:
        urlretrieve(LO_URL, temp_msi, reporthook)
    except Exception as e:
        parent.progress.stop()
        messagebox.showerror("下载失败", str(e))
        return False
    parent.progress.stop()
    parent.progress["value"] = 0
    cmd = ["msiexec", "/i", temp_msi, "/qn",
           "INSTALLDESKTOPSHORTCUT=0", "REBOOT=ReallySuppress"]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        messagebox.showerror("安装失败", proc.stderr)
        return False
    return True

# -------------------- 线程安全进度桥 + 百分比标签 --------------------
class ProgressBridge:
    def __init__(self, root, bar: ttk.Progressbar, label: tk.Label):
        self.root = root
        self.bar = bar
        self.label = label
        self.queue = Queue()
        self.root.after(50, self._poll)

    def _poll(self):
        while not self.queue.empty():
            val = self.queue.get_nowait()
            self.bar["value"] = val
            self.label.config(text=f"{val} %")
        self.root.after(50, self._poll)

    def update(self, val: int):
        self.queue.put(min(val, 100))

# -------------------- GUI --------------------
class WordToPdfConverter:
    DEFAULT_OUT = os.path.join(os.path.expanduser("~"), "Desktop", "PDF转换结果")

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Office转PDF小工具  (MS/WPS/LO 自动适配)")
        self.root.geometry("600x250")
        self.root.resizable(False, False)
        self.word_path = ""
        self.pdf_path = ""
        os.makedirs(self.DEFAULT_OUT, exist_ok=True)

        engine = which_office()
        if engine == "None":
            ok = messagebox.askyesno(
                "未检测到任何 Office 环境",
                "本机未找到 MS Office、WPS 或 LibreOffice。\n是否立即下载并安装 LibreOffice？",
            )
            if ok and download_install_lo(self):
                messagebox.showinfo("完成", "LibreOffice 安装成功！请重新启动本工具。")
                self.root.quit()
                return
            else:
                messagebox.showwarning("缺少依赖", "无法继续，程序即将退出。")
                self.root.destroy()
                return
        self.build_ui()

    def build_ui(self):
        tk.Label(self.root, text=f"Office转PDF小工具  （引擎：{which_office()}）", font=("微软雅黑", 14)).pack(pady=8)
        row1 = tk.Frame(self.root)
        row1.pack(fill="x", padx=20, pady=4)
        tk.Label(row1, text="Office文件：", width=10, anchor="w").pack(side="left")
        self.word_entry = tk.Entry(row1, width=45)
        self.word_entry.pack(side="left", padx=(0, 5))
        tk.Button(row1, text="浏览...", width=8, command=self.select_word).pack(side="left")
        row2 = tk.Frame(self.root)
        row2.pack(fill="x", padx=20, pady=4)
        tk.Label(row2, text="输出目录：", width=10, anchor="w").pack(side="left")
        self.out_entry = tk.Entry(row2, width=45)
        self.out_entry.pack(side="left", padx=(0, 5))
        self.out_entry.insert(0, self.DEFAULT_OUT)
        tk.Button(row2, text="浏览...", width=8, command=self.select_out).pack(side="left")

        # 进度条 + 百分比
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("green.Horizontal.TProgressbar", foreground='green', background='green')
        self.progress = ttk.Progressbar(self.root, mode="determinate", length=540,
                                        style="green.Horizontal.TProgressbar")
        self.progress.pack(pady=4)
        self.percent = tk.Label(self.root, text="0 %")
        self.percent.pack()
        self.bridge = ProgressBridge(self.root, self.progress, self.percent)

        row3 = tk.Frame(self.root)
        row3.pack(fill="x", padx=20, pady=6)
        self.convert_btn = tk.Button(row3, text="开始转换", bg="#4CAF50", fg="white", width=15,
                                     command=self.start_convert)
        self.convert_btn.pack(side="left")
        self.open_btn = tk.Button(row3, text="打开文件", width=15, state="disabled", command=self.open_pdf)
        self.open_btn.pack(side="right")

    def select_word(self):
        f = filedialog.askopenfilename(
            filetypes=[("Office 文件", "*.docx *.doc *.xlsx *.xls *.pptx *.ppt")]
        )
        if f:
            self.word_path = f
            self.word_entry.delete(0, tk.END)
            self.word_entry.insert(0, f)

    def select_out(self):
        d = filedialog.askdirectory()
        if d:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, d)

    # -------------------- 按钮触发 --------------------
    def start_convert(self):
        if not self.word_path:
            messagebox.showwarning("提示", "请先选择Office文件！")
            return
        self.convert_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self._thread_worker, daemon=True).start()

    def _thread_worker(self):
        pythoncom.CoInitialize()
        try:
            out_dir = self.out_entry.get()
            pdf_path, err = _do_convert(self.word_path, out_dir, self.bridge)
            if err:
                self.root.after(0, lambda: messagebox.showerror("错误", err))
            else:
                self.pdf_path = pdf_path
                self.root.after(0, lambda: messagebox.showinfo("完成", "PDF已生成！"))
                self.root.after(0, lambda: self.open_btn.config(state="normal"))
        finally:
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.convert_btn.config(state="normal"))

    def open_pdf(self):
        if os.path.isfile(self.pdf_path):
            os.startfile(self.pdf_path) if sys.platform == "win32" else subprocess.call(
                ["open", self.pdf_path] if sys.platform == "darwin" else ["xdg-open", self.pdf_path])

# -------------------- 启动 --------------------
if __name__ == "__main__":
    root = tk.Tk()
    WordToPdfConverter(root)
    root.mainloop()