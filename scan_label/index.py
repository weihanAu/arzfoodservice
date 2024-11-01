import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
import os
import csv_generater
import print_model
import pandas as pd
from datetime import datetime

def main():
    # 创建一个Tkinter窗口
    root = tk.Tk()
    root.title("Order Number Entry")
    root.geometry("400x300")

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
            # messagebox.showinfo("Order Number", f"Your order number is: {user_input}")
            try:
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                label_file_name=user_input+'_'+current_time+'_label.pdf'
                print_model.create_label_pdf(user_input,label_file_name)
                entry.delete(0, tk.END)  # 清空输入框
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
                    update_csv_range()
                except Exception as e:
                    messagebox.showerror("Error", str(e))
            else:
                messagebox.showwarning("Invalid File", "Please select a valid Excel (.xlsx) file.")

    # 创建一个加载 Excel 文件的按钮
    load_button = tk.Button(root, text="Load Excel File", command=load_excel_file)
    load_button.pack(pady=10)

    # read the start and end order number
    # Read CSV file
    def update_csv_range(event=None):
        total_orders=0
        distinct_sales_orders=0
        try:
            # 尝试读取 CSV 文件
            df = pd.read_csv('orders_labels.csv')
            
            # 获取第一个非空的 'SalesOrder.Number' 值
            number1 = df['SalesOrder.Number'].dropna().iloc[0]  # 从开头获取第一个匹配值
            # 获取最后一个非空的 'SalesOrder.Number' 值
            number2 = df['SalesOrder.Number'].dropna().iloc[-1]  # 从末尾获取第一个匹配值
           
            # 检查 number1 和 number2 是否存在
            if number1 and number2:
                result = f"{number1} - {number2}"
                # 统计订单总数
                total_orders = df['SalesOrder.Number'].dropna().count()
                # return how many orders
                distinct_sales_orders = df['SalesOrder.Number'].nunique()
            else:
                result = 'No valid data found'
                
        except FileNotFoundError:
            result = 'No CSV file found!'
        except Exception as e:
            result = f"Error: {str(e)}"
        
        # 更新标签的文本
        result_label.config(text=f"Current CSV range: {result}")
        result_label_count.config(text=f"Total labels: {total_orders}")
        result_label_count_orders.config(text=f"Total Orders: {distinct_sales_orders}")
   
    # 创建一个标签用于显示结果
    result_label = tk.Label(root, text="Current CSV range: ")
    result_label.pack(pady=5) 
    result_label_count_orders = tk.Label(root, text="Total orders: ")
    result_label_count_orders.pack(pady=5)
    result_label_count = tk.Label(root, text="Total labels: ")
    result_label_count.pack(pady=5)

    update_csv_range()

    # 保持主窗口打开
    root.mainloop()
    
if __name__ == "__main__":
    main()
