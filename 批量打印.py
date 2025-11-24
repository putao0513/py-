import tkinter as tk
from tkinter import filedialog, messagebox
import win32print
import win32api
import os

# ------------------------
# 静默打印一个文件（PDF/图片都行）
# ------------------------
def silent_print(file_path, printer_name=None):
    if printer_name is None:
        printer_name = win32print.GetDefaultPrinter()

    # 使用 Windows ShellExecute 的 print 命令，无窗口、无弹窗
    try:
        win32api.ShellExecute(
            0,
            "print",
            file_path,
            f'/d:"{printer_name}"',
            ".",
            0  # 0 = SW_HIDE 隐藏窗口
        )
    except Exception as e:
        print(f"打印失败: {file_path}, 错误: {e}")


# ------------------------
#  提交批量打印
# ------------------------
def start_batch_print():
    if not files:
        messagebox.showwarning("提示", "请先选择文件")
        return

    printer = win32print.GetDefaultPrinter()

    for f in files:
        silent_print(f, printer)

    messagebox.showinfo("完成", f"已经提交打印任务，共 {len(files)} 个文件。")


# ------------------------
# GUI 部分
# ------------------------
def select_files():
    global files
    files = filedialog.askopenfilenames(
        title="选择图片或PDF",
        filetypes=[
            ("图片/PDF", "*.pdf;*.jpg;*.jpeg;*.png;*.bmp"),
            ("所有文件", "*.*")
        ]
    )
    file_list.delete(0, tk.END)
    for f in files:
        file_list.insert(tk.END, f)


# ------------------------
# 主界面
# ------------------------
root = tk.Tk()
root.title("批量静默打印工具")
root.geometry("600x400")

files = []

# 文件列表
file_list = tk.Listbox(root, width=80, height=15)
file_list.pack(pady=10)

btn_frame = tk.Frame(root)
btn_frame.pack()

tk.Button(btn_frame, text="选择文件", width=15, command=select_files).grid(row=0, column=0, padx=10)
tk.Button(btn_frame, text="开始打印", width=15, command=start_batch_print).grid(row=0, column=1, padx=10)

root.mainloop()


//gemini版本
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import win32api
import win32print
import time
import os

class BatchPrintApp:
    def __init__(self, root):
        self.root = root
        self.root.title("极速批量打印工具")
        self.root.geometry("600x450")
        
        # 存储选中的文件路径
        self.files_to_print = []

        # --- 界面布局 ---
        
        # 1. 顶部说明
        self.lbl_info = tk.Label(root, text="已选文件列表 (支持 PDF, JPG, PNG 等):", font=("Arial", 10))
        self.lbl_info.pack(pady=5, anchor="w", padx=10)

        # 2. 文件列表展示区
        self.txt_list = scrolledtext.ScrolledText(root, height=15)
        self.txt_list.pack(fill="both", expand=True, padx=10, pady=5)

        # 3. 底部按钮区
        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(pady=10, fill="x", padx=10)

        self.btn_add = tk.Button(self.btn_frame, text="添加文件", command=self.add_files, bg="#e1f5fe", height=2)
        self.btn_add.pack(side="left", fill="x", expand=True, padx=5)

        self.btn_clear = tk.Button(self.btn_frame, text="清空列表", command=self.clear_files, height=2)
        self.btn_clear.pack(side="left", fill="x", expand=True, padx=5)

        self.btn_print = tk.Button(self.btn_frame, text="开始批量打印", command=self.start_printing, bg="#4caf50", fg="white", font=("Arial", 10, "bold"), height=2)
        self.btn_print.pack(side="left", fill="x", expand=True, padx=5)

        # 4. 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪 - 将使用系统默认打印机")
        self.lbl_status = tk.Label(root, textvariable=self.status_var, bd=1, relief="sunken", anchor="w")
        self.lbl_status.pack(side="bottom", fill="x")

    def add_files(self):
        """打开文件选择框"""
        filetypes = [
            ("支持的文件", "*.pdf;*.jpg;*.jpeg;*.png;*.bmp"),
            ("PDF 文档", "*.pdf"),
            ("图片文件", "*.jpg;*.jpeg;*.png;*.bmp"),
            ("所有文件", "*.*")
        ]
        files = filedialog.askopenfilenames(title="选择要打印的文件", filetypes=filetypes)
        
        if files:
            for f in files:
                # 避免重复添加
                if f not in self.files_to_print:
                    self.files_to_print.append(f)
                    self.txt_list.insert(tk.END, f + "\n")
            self.status_var.set(f"当前共有 {len(self.files_to_print)} 个文件等待打印。")

    def clear_files(self):
        """清空列表"""
        self.files_to_print = []
        self.txt_list.delete(1.0, tk.END)
        self.status_var.set("列表已清空。")

    def start_printing(self):
        """执行打印逻辑"""
        if not self.files_to_print:
            messagebox.showwarning("提示", "请先添加文件！")
            return

        # 获取默认打印机名称，用于显示
        try:
            default_printer = win32print.GetDefaultPrinter()
            confirm = messagebox.askyesno("确认打印", f"将发送 {len(self.files_to_print)} 个文件到默认打印机：\n\n【{default_printer}】\n\n是否继续？")
            if not confirm:
                return
        except Exception as e:
            messagebox.showerror("错误", f"无法获取打印机信息: {str(e)}")
            return

        # 遍历打印
        success_count = 0
        for index, file_path in enumerate(self.files_to_print):
            try:
                self.status_var.set(f"正在发送第 {index + 1}/{len(self.files_to_print)} 个文件: {os.path.basename(file_path)}")
                self.root.update() # 刷新界面显示

                # --- 核心打印代码 ---
                # ShellExecute(hwnd, operation, file, params, dir, show_cmd)
                # operation="print": 执行打印命令
                # show_cmd=0: 尝试隐藏窗口 (SW_HIDE)
                win32api.ShellExecute(0, "print", file_path, None, ".", 0)
                
                success_count += 1
                
                # *重要*：添加延时，防止发送太快导致打印机队列卡死或程序崩溃
                time.sleep(2) 

            except Exception as e:
                print(f"打印失败: {file_path}, 错误: {e}")
        
        self.status_var.set(f"任务完成。成功发送 {success_count} 个任务。")
        messagebox.showinfo("完成", "所有打印任务已发送至打印机队列。")

if __name__ == "__main__":
    root = tk.Tk()
    app = BatchPrintApp(root)
    root.mainloop()
