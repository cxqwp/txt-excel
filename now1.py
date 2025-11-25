# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
from setuptools import setup
setup( 
    name='my_project', 
    version='0.1', 
    author='cxpwp', 
    author_email='SincereDreamswithJade-likeIntegrity@outlook.com', 
    description='将TXT文本文件转换为Excel格式，支持数据编辑和自定义分隔符', 
    packages=['my_package'],)
class TxtToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("TXT转Excel工具")
        self.root.geometry("1000x600")
        
        # 数据存储
        self.data = []
        self.headers = []
        self.txt_file_path = ""
        
        # 创建界面
        self.setup_ui()
        
    def setup_ui(self):
        # 顶部工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        # 文件操作按钮
        ttk.Button(toolbar, text="打开TXT文件", command=self.load_txt_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="保存为Excel", command=self.save_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="添加行", command=self.add_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="删除选中行", command=self.delete_selected_rows).pack(side=tk.LEFT, padx=5)
        
        # 分隔符选择
        ttk.Label(toolbar, text="分隔符:").pack(side=tk.LEFT, padx=5)
        self.delimiter_var = tk.StringVar(value="|")
        delimiter_combo = ttk.Combobox(toolbar, textvariable=self.delimiter_var, 
                                      values=["|", ",", ";", "\t", " "], width=5)
        delimiter_combo.pack(side=tk.LEFT, padx=5)
        delimiter_combo.bind("<<ComboboxSelected>>", self.on_delimiter_change)
        
        # 添加列操作按钮
        ttk.Button(toolbar, text="添加列", command=self.add_column).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="编辑列名", command=lambda: self.edit_column_name(0)).pack(side=tk.LEFT, padx=5)
        
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建Treeview用于显示和编辑数据
        self.tree_frame = ttk.Frame(main_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建滚动条
        tree_scroll = ttk.Scrollbar(self.tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scroll = ttk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建Treeview
        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=tree_scroll.set, 
                                xscrollcommand=x_scroll.set, selectmode='extended')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        tree_scroll.config(command=self.tree.yview)
        x_scroll.config(command=self.tree.xview)
        
        # 绑定双击编辑事件
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # 添加右键菜单用于编辑列名
        self.tree.bind("<Button-3>", self.on_right_click)
        
        # 添加右键菜单用于编辑列名
        self.tree.bind("<Button-3>", self.on_right_click)
        
        # 状态栏
        self.status_bar = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def load_txt_file(self):
        """加载TXT文件"""
        file_path = filedialog.askopenfilename(
            title="选择TXT文件",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            self.txt_file_path = file_path
            delimiter = self.delimiter_var.get()
            
            # 读取文件
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()
            
            if not lines:
                messagebox.showwarning("警告", "文件为空")
                return
            
            # 解析数据
            self.data = []
            self.headers = []
            
            # 假设第一行是标题行
            first_line = lines[0].strip()
            self.headers = first_line.split(delimiter)
            
            # 解析数据行
            for line in lines[1:]:
                line = line.strip()
                if line:
                    row_data = line.split(delimiter)
                    # 确保每行数据与标题列数一致
                    while len(row_data) < len(self.headers):
                        row_data.append("")
                    self.data.append(row_data[:len(self.headers)])
            
            # 更新界面
            self.update_treeview()
            self.status_bar.config(text=f"已加载文件: {os.path.basename(file_path)} (共{len(self.data)}行数据)")
            
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def update_treeview(self):
        """更新Treeview显示"""
        # 清除现有内容
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 设置列
        if self.headers:
            self.tree["columns"] = self.headers
            self.tree.heading("#0", text="行号")
            self.tree.column("#0", width=50, stretch=tk.NO)
            
            for header in self.headers:
                self.tree.heading(header, text=header)
                self.tree.column(header, width=150, minwidth=50)
        
        # 插入数据
        for i, row in enumerate(self.data, 1):
            self.tree.insert("", tk.END, text=str(i), values=row)
    
    def on_delimiter_change(self, event=None):
        """分隔符改变时重新加载文件"""
        if self.txt_file_path:
            self.load_txt_file()
    
    def on_double_click(self, event):
        """双击编辑单元格"""
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        
        if column == "#0":  # 行号列
            return
            
        column_index = int(column.replace("#", "")) - 1
        old_value = self.tree.item(item, "values")[column_index]
        
        # 创建编辑框
        x, y, width, height = self.tree.bbox(item, column)
        edit_window = tk.Toplevel(self.root)
        edit_window.geometry(f"{width}x{height}+{x}+{y}")
        edit_window.overrideredirect(True)
        
        edit_var = tk.StringVar(value=old_value)
        edit_entry = ttk.Entry(edit_window, textvariable=edit_var)
        edit_entry.pack(fill=tk.BOTH, expand=True)
        edit_entry.focus()
        edit_entry.select_range(0, tk.END)
        
        def save_edit():
            new_value = edit_var.get()
            values = list(self.tree.item(item, "values"))
            values[column_index] = new_value
            self.tree.item(item, values=values)
            
            # 更新数据
            row_index = int(self.tree.item(item, "text")) - 1
            self.data[row_index] = values
            
            edit_window.destroy()
        
        def cancel_edit():
            edit_window.destroy()
        
        edit_entry.bind("<Return>", lambda e: save_edit())
        edit_entry.bind("<Escape>", lambda e: cancel_edit())
        edit_window.bind("<FocusOut>", lambda e: save_edit())
    
    def add_row(self):
        """添加新行"""
        if not self.headers:
            # 如果没有表头，创建默认表头
            self.headers = ["列1", "列2", "列3"]
        
        # 创建新行数据
        new_row = [""] * len(self.headers)
        self.data.append(new_row)
        
        # 插入到Treeview
        row_num = len(self.data)
        self.tree.insert("", tk.END, text=str(row_num), values=new_row)
        
        self.status_bar.config(text=f"已添加新行 (共{len(self.data)}行数据)")
    
    def delete_selected_rows(self):
        """删除选中行"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的行")
            return
        
        if messagebox.askyesno("确认", f"确定要删除选中的{len(selected_items)}行吗？"):
            # 获取要删除的行号（从大到小排序）
            rows_to_delete = []
            for item in selected_items:
                row_num = int(self.tree.item(item, "text")) - 1
                rows_to_delete.append(row_num)
            
            rows_to_delete.sort(reverse=True)
            
            # 从数据列表中删除
            for row_num in rows_to_delete:
                del self.data[row_num]
            
            # 重新加载显示
            self.update_treeview()
            self.status_bar.config(text=f"已删除选中行 (剩余{len(self.data)}行数据)")
    
    def on_right_click(self, event):
        """右键点击事件"""
        # 获取点击的列
        column = self.tree.identify_column(event.x)
        
        if column == "#0":  # 行号列
            return
            
        column_index = int(column.replace("#", "")) - 1
        
        # 创建右键菜单
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="编辑列名", command=lambda: self.edit_column_name(column_index))
        menu.add_command(label="添加列", command=lambda: self.add_column(column_index))
        menu.add_command(label="删除列", command=lambda: self.delete_column(column_index))
        
        # 显示菜单
        menu.post(event.x_root, event.y_root)
    
    def edit_column_name(self, column_index):
        """编辑列名"""
        if not self.headers:
            return
            
        old_name = self.headers[column_index]
        
        # 创建编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑列名")
        dialog.geometry("300x120")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text=f"原列名: {old_name}").pack(pady=10)
        
        name_var = tk.StringVar(value=old_name)
        entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        entry.pack(pady=5)
        entry.focus()
        entry.select_range(0, tk.END)
        
        def save_name():
            new_name = name_var.get().strip()
            if new_name and new_name != old_name:
                # 更新列名
                self.headers[column_index] = new_name
                # 重新加载Treeview
                self.update_treeview()
                self.status_bar.config(text=f"列名已修改: {old_name} → {new_name}")
            dialog.destroy()
        
        def cancel():
            dialog.destroy()
        
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="确定", command=save_name).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=cancel).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        entry.bind("<Return>", lambda e: save_name())
        entry.bind("<Escape>", lambda e: cancel())
    
    def add_column(self, insert_index=None):
        """添加新列"""
        if insert_index is None:
            insert_index = len(self.headers)
        
        # 创建编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("添加列")
        dialog.geometry("300x120")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="新列名:").pack(pady=10)
        
        name_var = tk.StringVar(value="新列")
        entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        entry.pack(pady=5)
        entry.focus()
        entry.select_range(0, tk.END)
        
        def save_name():
            new_name = name_var.get().strip()
            if new_name:
                # 插入新列名
                self.headers.insert(insert_index + 1, new_name)
                
                # 为所有数据行添加新列的空值
                for i in range(len(self.data)):
                    self.data[i].insert(insert_index + 1, "")
                
                # 重新加载Treeview
                self.update_treeview()
                self.status_bar.config(text=f"已添加新列: {new_name}")
            dialog.destroy()
        
        def cancel():
            dialog.destroy()
        
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="确定", command=save_name).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=cancel).pack(side=tk.LEFT, padx=5)
        
        entry.bind("<Return>", lambda e: save_name())
        entry.bind("<Escape>", lambda e: cancel())
    
    def delete_column(self, column_index):
        """删除列"""
        if len(self.headers) <= 1:
            messagebox.showwarning("警告", "至少需要保留一列")
            return
        
        column_name = self.headers[column_index]
        
        if messagebox.askyesno("确认", f"确定要删除列 '{column_name}' 吗？"):
            # 删除列名
            del self.headers[column_index]
            
            # 删除所有数据行中的对应列
            for i in range(len(self.data)):
                if column_index < len(self.data[i]):
                    del self.data[i][column_index]
            
            # 重新加载Treeview
            self.update_treeview()
            self.status_bar.config(text=f"已删除列: {column_name}")
    
    def save_to_excel(self):
        """保存为Excel文件"""
        if not self.data:
            messagebox.showwarning("警告", "没有数据可保存")
            return
        
        # 生成默认文件名
        default_name = f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        file_path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=default_name
        )
        
        if not file_path:
            return
        
        try:
            # 创建DataFrame，尝试将数字列转换为数值类型
            df = pd.DataFrame(self.data, columns=self.headers)
            
            # 尝试将可以转换为数字的列转换为数值类型
            for col in df.columns:
                # 尝试将列转换为数值类型
                try:
                    # 先尝试转换为float，如果成功再检查是否可以转换为int
                    numeric_series = pd.to_numeric(df[col], errors='coerce')
                    # 检查转换后的非空值比例，如果大部分都能转换，则使用数值类型
                    non_null_ratio = numeric_series.notna().sum() / len(numeric_series)
                    if non_null_ratio > 0.8:  # 80%以上的值可以转换为数字
                        df[col] = numeric_series
                        # 如果是整数且没有小数部分，转换为int类型
                        if numeric_series.dropna().apply(lambda x: x == int(x)).all():
                            df[col] = numeric_series.astype('Int64')  # 使用可空整数类型
                except:
                    # 转换失败，保持原样（文本格式）
                    pass
            
            # 创建Excel写入器
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='数据', index=False)
                
                # 获取工作簿和工作表
                workbook = writer.book
                worksheet = writer.sheets['数据']
                
                # 设置列宽
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                
                # 设置标题行格式
                from openpyxl.styles import Font, Alignment
                
                for cell in worksheet[1]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
            
            messagebox.showinfo("成功", f"文件已保存到:\n{file_path}")
            self.status_bar.config(text=f"已保存Excel文件: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存文件失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TxtToExcelConverter(root)
    root.mainloop()
setup( 
    name='txt转excel', 
    version='0.1', 
    author='cxpwp', 
    author_email='SincereDreamswithJade-likeIntegrity@outlook.com', 
    packages=['my_package'],)
