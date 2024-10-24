import tkinter as tk
from tkinter import simpledialog

def main():
    # 创建一个Tkinter窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 弹出输入对话框
    user_input = simpledialog.askstring(title="order number", prompt="please scan or enter order number:")

    if user_input:
        # 创建一个新窗口来显示用户的输入
        result_window = tk.Toplevel(root)
        result_window.title("order number")

        # 在新窗口中显示用户输入的字
        label = tk.Label(result_window, text=f"order number: {user_input}", font=("Arial", 16))
        label.pack(padx=20, pady=20)

        # 保持窗口打开
        result_window.mainloop()

if __name__ == "__main__":
    main()
