import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
import os
import csv_generater
import print_model

def main():
    # 创建一个Tkinter窗口
    root = tk.Tk()
    root.title("Order Number Entry")
    root.geometry("300x150")

    # 创建一个Label用于说明
    label = tk.Label(root, text="Please enter your order number:")
    label.pack(pady=10)

    # 创建一个输入框（Entry）
    entry = tk.Entry(root, width=30)
    entry.pack(pady=5)

    def submit_order_number():
        user_input = entry.get()
        if user_input:
            # 显示用户输入的订单号
            messagebox.showinfo("Order Number", f"Your order number is: {user_input}")
            entry.delete(0, tk.END)  # 清空输入框
            try:
                print_model.create_label_pdf(111111,"label_output.pdf")
            except Exception as e:
                    messagebox.showerror("Error", str(e))

    # 监听回车键，用户按下回车后会自动提交输入内容
    root.bind('<Return>', lambda event: submit_order_number())

    # 创建一个提交按钮
    submit_button = tk.Button(root, text="Print Order Labels", command=submit_order_number)
    submit_button.pack(pady=10)

    # create a new button to load files in.
    def load_excel_file():
        # 弹出文件选择器，让用户选择 .xlsx 文件
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            # 检查文件是否是 .xlsx 文件
            if file_path.endswith('.xlsx'):
                messagebox.showinfo("File Loaded", f"Loaded file: {os.path.basename(file_path)}")
                try: 
                    csv_generater.duplicate_rows_based_on_quantity(file_path)
                    messagebox.showinfo("Success", "Labels generated successfully!")
                except Exception as e:
                    messagebox.showerror("Error", str(e))
            else:
                messagebox.showwarning("Invalid File", "Please select a valid Excel (.xlsx) file.")

    # 创建一个加载 Excel 文件的按钮
    load_button = tk.Button(root, text="Load Excel File", command=load_excel_file)
    load_button.pack(pady=10)

    # 保持主窗口打开
    root.mainloop()

   

    
if __name__ == "__main__":
    main()
