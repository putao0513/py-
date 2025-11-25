import os
import sys  # 必须导入 sys
import time
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

# ================= ⚙️ 核心配置修改 =================

def get_resource_path(relative_path):
    """
    获取资源文件的绝对路径。
    用于兼容 开发环境 和 打包后的 exe 环境。
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的临时目录
        return os.path.join(sys._MEIPASS, relative_path)
    # 正常开发环境的目录
    return os.path.join(os.path.abspath("."), relative_path)

# 动态获取打包在内部的 SumatraPDF 路径
SUMATRA_PATH = get_resource_path("SumatraPDF.exe")
PRINT_DELAY = 2 

# ===================================================

class PrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量打印小助手 (便携版)")
        self.root.geometry("600x450")

        # 1. 顶部操作区
        top_frame = tk.Frame(root, pady=10)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        self.btn_select = tk.Button(top_frame, text="📂 选择文件并打印", font=("微软雅黑", 12, "bold"), 
                                    bg="#4CAF50", fg="white", height=2, width=20,
                                    command=self.start_process_logic)
        self.btn_select.pack()

        self.lbl_status = tk.Label(top_frame, text="准备就绪 (内置打印引擎)", fg="gray")
        self.lbl_status.pack(pady=5)

        self.log_area = scrolledtext.ScrolledText(root, state='disabled', height=15, font=("Consolas", 10))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log("欢迎使用！本程序已内置 PDF 打印引擎，无需安装额外软件。")

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def start_process_logic(self):
        files = filedialog.askopenfilenames(
            title="请选择要打印的文件",
            filetypes=[("支持的文件", "*.pdf *.jpg *.jpeg *.png *.bmp"), ("PDF 文件", "*.pdf"), ("图片文件", "*.jpg;*.png")]
        )

        if not files: return

        count = len(files)
        if not messagebox.askyesno("打印确认", f"选中 {count} 个文件。\n是否立即打印？"):
            self.log("🚫 操作已取消")
            return

        self.btn_select.config(state=tk.DISABLED, bg="gray", text="正在处理...")
        
        thread = threading.Thread(target=self.process_files, args=(files,))
        thread.daemon = True
        thread.start()

    def process_files(self, files):
        total = len(files)
        success_count = 0
        self.log("-" * 40)
        
        for index, file_path in enumerate(files, 1):
            filename = os.path.basename(file_path)
            self.lbl_status.config(text=f"正在打印 ({index}/{total}): {filename}")
            
            try:
                self.print_single_file(file_path)
                success_count += 1
            except Exception as e:
                self.log(f"❌ [失败] {filename}: {str(e)}")

            if index < total:
                time.sleep(PRINT_DELAY)

        self.log("-" * 40 + "\n🎉 任务完成！")
        self.lbl_status.config(text="任务完成")
        self.root.after(0, lambda: self.btn_select.config(state=tk.NORMAL, bg="#4CAF50", text="📂 选择文件并打印"))
        messagebox.showinfo("完成", f"打印结束！成功 {success_count} 个。")

    def print_single_file(self, file_path):
        abs_path = os.path.abspath(file_path)
        ext = os.path.splitext(abs_path)[1].lower()

        # 检查内置工具是否存在
        if ext == '.pdf' and not os.path.exists(SUMATRA_PATH):
            raise Exception("内置打印组件丢失")

        if ext == '.pdf':
            subprocess.run([SUMATRA_PATH, "-print-to-default", "-exit-on-print", abs_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.log(f"✅ [PDF] {os.path.basename(abs_path)}")
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            subprocess.run(["mspaint", "/p", abs_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.log(f"✅ [IMG] {os.path.basename(abs_path)}")
        else:
            os.startfile(abs_path, "print")

if __name__ == "__main__":
    root = tk.Tk()
    app = PrinterApp(root)
    root.mainloop()
