import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re

class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 拆分工具")
        self.root.geometry("400x150") # 设置窗口大小

        self.file_path = ""

        # 创建主框架
        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择部分
        select_frame = tk.Frame(main_frame)
        select_frame.pack(fill=tk.X, pady=5)
        
        self.select_button = tk.Button(select_frame, text="选择待拆分 Excel", command=self.select_file)
        self.select_button.pack(side=tk.LEFT, padx=(0, 10))

        self.file_label = tk.Label(select_frame, text="尚未选择文件", fg="grey", anchor="w")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 拆分按钮部分
        self.split_button = tk.Button(main_frame, text="开始拆分", command=self.split_excel, state=tk.DISABLED)
        self.split_button.pack(pady=10, fill=tk.X)

    def select_file(self):
        """
        打开文件选择对话框。
        """
        path = filedialog.askopenfilename(
            title="请选择一个Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.file_path = path
            # 使用os.path.basename获取文件名
            filename = os.path.basename(path)
            self.file_label.config(text=filename, fg="black")
            self.split_button.config(state=tk.NORMAL) # 启用拆分按钮

    def clean_filename(self, name):
        """
        移除文件名中的非法字符。
        """
        if not isinstance(name, str):
            name = str(name)
        return re.sub(r'[\\/*?:"<>|]', '_', name)

    def split_excel(self):
        """
        根据第一列的内容将Excel文件拆分成多个文件。
        """
        if not self.file_path:
            messagebox.showwarning("警告", "请先选择一个Excel文件。")
            return

        try:
            df = pd.read_excel(self.file_path)

            if df.empty:
                messagebox.showwarning("警告", "选择的Excel文件是空的。")
                return

            first_column_name = df.columns[0]
            unique_values = df[first_column_name].unique()
            output_dir = os.path.dirname(self.file_path)

            for value in unique_values:
                df_split = df[df[first_column_name] == value]
                clean_name = self.clean_filename(value)
                new_file_name = f"{clean_name}.xlsx"
                new_file_path = os.path.join(output_dir, new_file_name)
                df_split.to_excel(new_file_path, index=False)

            messagebox.showinfo("成功", f"处理完成！\n文件已成功拆分并保存在以下目录：\n{output_dir}")

        except Exception as e:
            messagebox.showerror("错误", f"处理过程中发生错误：\n{e}")
        finally:
            # 处理完成后可以重置界面状态
            self.file_path = ""
            self.file_label.config(text="尚未选择文件", fg="grey")
            self.split_button.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()