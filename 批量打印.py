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
