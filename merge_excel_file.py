import pandas as pd
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os


class ExcelMergerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel文件合并工具")
        self.root.geometry("700x500")
        self.root.resizable(True, True)

        self.folder_path = ""
        self.header_start_row = 1  # 默认标题行开始行
        self.header_end_row = 1    # 默认标题行结束行
        self.setup_ui()

    def setup_ui(self):
        # 主标题
        title_label = tk.Label(self.root, text="Excel文件合并工具", font=("Microsoft YaHei", 16, "bold"))
        title_label.pack(pady=20)

        # 说明文字
        instruction_label = tk.Label(self.root, text="请选择包含Excel文件的文件夹，程序将自动合并所有Excel文件",
                                     font=("Microsoft YaHei", 10), wraplength=600)
        instruction_label.pack(pady=10)

        # 标题行设置框架
        header_frame = tk.LabelFrame(self.root, text="标题行设置", font=("Microsoft YaHei", 10, "bold"), padx=20, pady=15)
        header_frame.pack(pady=15, padx=20, fill="x")

        # 标题行说明
        header_instruction = tk.Label(header_frame, 
                                    text="请设置标题行的范围（例如：第1行到第3行作为标题行）\n注意：所有Excel文件的标题行必须完全一致才能成功合并\n提示：如果不选择标题行，程序将默认使用第1行作为标题行",
                                    font=("Microsoft YaHei", 9), 
                                    fg="blue",
                                    wraplength=500,
                                    justify="left")
        header_instruction.pack(pady=(0, 15))

        # 标题行输入框架
        header_input_frame = tk.Frame(header_frame)
        header_input_frame.pack(fill="x")

        # 开始行标签和输入框
        start_label = tk.Label(header_input_frame, text="标题行开始行号:", font=("Microsoft YaHei", 10))
        start_label.pack(side="left", padx=(0, 10))
        
        self.start_entry = tk.Entry(header_input_frame, width=8, font=("Microsoft YaHei", 10))
        self.start_entry.insert(0, "1")
        self.start_entry.pack(side="left", padx=(0, 20))

        # 结束行标签和输入框
        end_label = tk.Label(header_input_frame, text="标题行结束行号:", font=("Microsoft YaHei", 10))
        end_label.pack(side="left", padx=(0, 10))
        
        self.end_entry = tk.Entry(header_input_frame, width=8, font=("Microsoft YaHei", 10))
        self.end_entry.insert(0, "1")
        self.end_entry.pack(side="left")

        # 文件夹选择框架
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(pady=20, padx=20, fill="x")

        # 选择文件夹按钮
        self.select_button = tk.Button(folder_frame, text="选择文件夹",
                                       command=self.select_folder,
                                       font=("Microsoft YaHei", 12),
                                       bg="#4CAF50", fg="white",
                                       relief="raised", padx=20, pady=10)
        self.select_button.pack(side="right", padx=(0, 10))

        # 显示选择的文件夹路径
        self.folder_label = tk.Label(folder_frame, text="未选择文件夹",
                                     font=("Microsoft YaHei", 10),
                                     fg="gray",
                                     wraplength=400,
                                     anchor="w")
        self.folder_label.pack(side="left", fill="x", expand=True)

        # 合并按钮
        self.merge_button = tk.Button(self.root, text="开始合并",
                                      command=self.merge_files,
                                      font=("Microsoft YaHei", 14, "bold"),
                                      bg="#2196F3", fg="white",
                                      relief="raised", padx=30, pady=15,
                                      state="disabled")
        self.merge_button.pack(pady=30)

        # 进度条
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(pady=10, padx=20, fill="x")

        # 状态标签
        self.status_label = tk.Label(self.root, text="", font=("Microsoft YaHei", 10))
        self.status_label.pack(pady=10)

    def validate_header_rows(self):
        """验证标题行输入"""
        try:
            start = int(self.start_entry.get())
            end = int(self.end_entry.get())
            
            if start < 1 or end < 1:
                messagebox.showerror("输入错误", "行号必须大于0")
                return False
                
            if start > end:
                messagebox.showerror("输入错误", "开始行号不能大于结束行号")
                return False
                
            self.header_start_row = start
            self.header_end_row = end
            return True
            
        except ValueError:
            messagebox.showerror("输入错误", "请输入有效的数字")
            return False

    def select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory(title="选择包含Excel文件的文件夹")
        if folder:
            self.folder_path = folder
            self.folder_label.config(text=f"已选择: {folder}", fg="green")
            self.merge_button.config(state="normal")
            self.status_label.config(text="文件夹已选择，可以开始合并")
        else:
            self.status_label.config(text="未选择文件夹")

    def check_header_consistency(self, file_paths):
        """检查所有文件的标题行是否一致"""
        if not file_paths:
            return False, "没有找到Excel文件"
            
        # 读取第一个文件的标题行作为标准
        try:
            first_df = pd.read_excel(file_paths[0], nrows=self.header_end_row)
            standard_headers = first_df.iloc[self.header_start_row-1:self.header_end_row]
        except Exception as e:
            return False, f"读取第一个文件失败: {str(e)}"
            
        # 检查其他文件的标题行
        for i, file_path in enumerate(file_paths[1:], 1):
            try:
                df = pd.read_excel(file_path, nrows=self.header_end_row)
                current_headers = df.iloc[self.header_start_row-1:self.header_end_row]
                
                if not standard_headers.equals(current_headers):
                    return False, f"文件 {os.path.basename(file_path)} 的标题行与第一个文件不一致"
                    
            except Exception as e:
                return False, f"读取文件 {os.path.basename(file_path)} 失败: {str(e)}"
                
        return True, "标题行检查通过"

    def merge_files(self):
        """合并Excel文件"""
        if not self.folder_path:
            messagebox.showwarning("警告", "请先选择文件夹")
            return

        # 验证标题行输入
        if not self.validate_header_rows():
            return

        # 禁用按钮，显示进度
        self.merge_button.config(state="disabled")
        self.select_button.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="正在检查标题行一致性...")
        self.root.update()

        try:
            # 获取所有Excel文件路径
            file_paths = glob.glob(os.path.join(self.folder_path, "*.xls")) + glob.glob(
                os.path.join(self.folder_path, "*.xlsx"))

            if not file_paths:
                messagebox.showerror("错误", f"在选择的文件夹中没有找到Excel文件")
                return

            self.status_label.config(text=f"找到 {len(file_paths)} 个Excel文件，正在检查标题行一致性...")
            self.root.update()

            # 检查标题行一致性
            is_consistent, message = self.check_header_consistency(file_paths)
            if not is_consistent:
                messagebox.showerror("标题行不一致", f"标题行检查失败：{message}\n\n请确保所有Excel文件的标题行完全一致后再进行合并。")
                return

            self.status_label.config(text="标题行检查通过，正在读取和合并文件...")
            self.root.update()

            # 读取并合并文件
            all_data = []
            success_count = 0
            error_files = []

            for i, file in enumerate(file_paths):
                try:
                    # 读取文件，跳过标题行
                    df = pd.read_excel(file, skiprows=self.header_end_row)
                    all_data.append(df)
                    success_count += 1
                    self.status_label.config(text=f"正在读取文件 {i + 1}/{len(file_paths)}: {os.path.basename(file)}")
                    self.root.update()
                except Exception as e:
                    error_files.append(f"{os.path.basename(file)}: {str(e)}")

            # 显示读取结果
            if error_files:
                error_message = "以下文件读取失败：\n" + "\n".join(error_files)
                messagebox.showwarning("读取警告", error_message)

            if not all_data:
                messagebox.showerror("错误", "没有成功读取任何文件")
                return

            self.status_label.config(text="正在合并数据...")
            self.root.update()

            # 合并数据
            merged_df = pd.concat(all_data, ignore_index=True)

            # 读取第一个文件的标题行
            first_df = pd.read_excel(file_paths[0], nrows=self.header_end_row)
            headers = first_df.iloc[self.header_start_row-1:self.header_end_row]

            # 保存结果，包含标题行
            output_file = os.path.join(self.folder_path, "merged_file.xlsx")
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 先写入标题行
                headers.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
                # 再写入数据行
                merged_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=self.header_end_row)

            # 显示成功信息
            success_message = f"合并完成！\n\n输出文件：{output_file}\n标题行范围：第{self.header_start_row}行到第{self.header_end_row}行\n总共合并了 {len(merged_df)} 行数据\n成功读取了 {success_count} 个文件"
            if error_files:
                success_message += f"\n\n有 {len(error_files)} 个文件读取失败"

            messagebox.showinfo("合并成功", success_message)
            self.status_label.config(text="合并完成！")

        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}")
            self.status_label.config(text="处理失败")
        finally:
            # 恢复按钮状态，停止进度条
            self.merge_button.config(state="normal")
            self.select_button.config(state="normal")
            self.progress.stop()

    def run(self):
        """运行应用程序"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ExcelMergerApp()
    app.run()
