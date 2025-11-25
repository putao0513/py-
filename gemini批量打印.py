import os
import time
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

# ================= ⚙️ 配置区域 =================
# 请确保此路径正确
SUMATRA_PATH = r"C:\Users\admin\AppData\Local\SumatraPDF\SumatraPDF.exe"
PRINT_DELAY = 2  # 打印间隔时间（秒）
# ==============================================

class PrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量打印小助手 (安全版)")
        self.root.geometry("600x450")

        # 1. 顶部操作区
        top_frame = tk.Frame(root, pady=10)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        self.btn_select = tk.Button(top_frame, text="📂 选择文件并打印", font=("微软雅黑", 12, "bold"), 
                                    bg="#4CAF50", fg="white", height=2, width=20,
                                    command=self.start_process_logic)
        self.btn_select.pack()

        # 2. 状态显示区
        self.lbl_status = tk.Label(top_frame, text="准备就绪，等待选择...", fg="gray")
        self.lbl_status.pack(pady=5)

        # 3. 日志输出区
        self.log_area = scrolledtext.ScrolledText(root, state='disabled', height=15, font=("Consolas", 10))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log("欢迎使用！点击上方按钮选择文件。")

    def log(self, message):
        """向文本框添加日志"""
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def start_process_logic(self):
        """处理点击按钮后的逻辑：选择 -> 确认 -> 启动线程"""
        # 1. 弹出文件选择框
        files = filedialog.askopenfilenames(
            title="请选择要打印的文件（可多选）",
            filetypes=[("支持的文件", "*.pdf *.jpg *.jpeg *.png *.bmp"), ("PDF 文件", "*.pdf"), ("图片文件", "*.jpg;*.png")]
        )

        if not files:
            return # 用户取消了选择，什么也不做

        count = len(files)
        
        # 2. 【新增】弹出确认框
        confirm = messagebox.askyesno(
            title="打印确认", 
            message=f"您已选中 {count} 个文件。\n\n是否立即开始打印？"
        )

        if not confirm:
            self.log(f"🚫 操作已取消 (选中了 {count} 个文件但未打印)")
            return # 用户点击了“否”，停止后续操作

        # 3. 用户点击了“是”，启动后台线程开始干活
        self.btn_select.config(state=tk.DISABLED, bg="gray", text="正在打印中...")
        
        thread = threading.Thread(target=self.process_files, args=(files,))
        thread.daemon = True
        thread.start()

    def process_files(self, files):
        """实际的后台打印循环"""
        total = len(files)
        success_count = 0

        self.log("-" * 40)
        self.log(f"🚀 开始任务，共 {total} 个文件")

        for index, file_path in enumerate(files, 1):
            filename = os.path.basename(file_path)
            self.lbl_status.config(text=f"正在打印 ({index}/{total}): {filename}")
            
            try:
                self.print_single_file(file_path)
                success_count += 1
            except Exception as e:
                self.log(f"❌ [失败] {filename}: {str(e)}")

            # 打印间隔
            if index < total:
                time.sleep(PRINT_DELAY)

        self.log("-" * 40)
        self.log(f"🎉 任务完成！成功: {success_count} / 总数: {total}")
        self.lbl_status.config(text="任务完成")
        
        # 恢复按钮状态
        self.root.after(0, lambda: self.btn_select.config(state=tk.NORMAL, bg="#4CAF50", text="📂 选择文件并打印"))
        messagebox.showinfo("完成", f"打印结束！\n成功发送 {success_count} 个文件。")

    def print_single_file(self, file_path):
        """调用外部工具打印"""
        abs_path = os.path.abspath(file_path)
        ext = os.path.splitext(abs_path)[1].lower()
        filename = os.path.basename(abs_path)

        self.log(f"🖨️ 正在发送: {filename}")

        # --- PDF (SumatraPDF) ---
        if ext == '.pdf':
            if not os.path.exists(SUMATRA_PATH):
                raise Exception("找不到 SumatraPDF 路径配置")
            
            subprocess.run([SUMATRA_PATH, "-print-to-default", "-exit-on-print", abs_path], check=True)
            self.log(f"✅ [PDF] 发送成功")

        # --- 图片 (mspaint) ---
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            subprocess.run(["mspaint", "/p", abs_path], check=True)
            self.log(f"✅ [图片] 发送成功")

        # --- 其他 ---
        else:
            self.log(f"⚠️ [系统默认] 调用默认程序...")
            os.startfile(abs_path, "print")

if __name__ == "__main__":
    root = tk.Tk()
    app = PrinterApp(root)
    root.mainloop()
