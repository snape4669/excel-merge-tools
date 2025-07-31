import pandas as pd
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

class ExcelMergerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel文件合并工具")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        self.folder_path = ""
        self.setup_ui()
        
    def setup_ui(self):
        # 主标题
        title_label = tk.Label(self.root, text="Excel文件合并工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=20)
        
        # 说明文字
        instruction_label = tk.Label(self.root, text="请选择包含Excel文件的文件夹，程序将自动合并所有Excel文件", 
                                   font=("Arial", 10), wraplength=500)
        instruction_label.pack(pady=10)
        
        # 文件夹选择框架
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(pady=20, padx=20, fill="x")
        
        # 选择文件夹按钮
        self.select_button = tk.Button(folder_frame, text="选择文件夹", 
                                     command=self.select_folder, 
                                     font=("Arial", 12), 
                                     bg="#4CAF50", fg="white",
                                     relief="raised", padx=20, pady=10)
        self.select_button.pack(side="right", padx=(0, 10))
        
        # 显示选择的文件夹路径
        self.folder_label = tk.Label(folder_frame, text="未选择文件夹", 
                                   font=("Arial", 10), 
                                   fg="gray", 
                                   wraplength=400,
                                   anchor="w")
        self.folder_label.pack(side="left", fill="x", expand=True)
        
        # 合并按钮
        self.merge_button = tk.Button(self.root, text="开始合并", 
                                    command=self.merge_files, 
                                    font=("Arial", 14, "bold"), 
                                    bg="#2196F3", fg="white",
                                    relief="raised", padx=30, pady=15,
                                    state="disabled")
        self.merge_button.pack(pady=30)
        
        # 进度条
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(pady=10, padx=20, fill="x")
        
        # 状态标签
        self.status_label = tk.Label(self.root, text="", font=("Arial", 10))
        self.status_label.pack(pady=10)
        
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
    
    def merge_files(self):
        """合并Excel文件"""
        if not self.folder_path:
            messagebox.showwarning("警告", "请先选择文件夹")
            return
            
        # 禁用按钮，显示进度
        self.merge_button.config(state="disabled")
        self.select_button.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="正在处理文件...")
        self.root.update()
        
        try:
            # 获取所有Excel文件路径
            file_paths = glob.glob(os.path.join(self.folder_path, "*.xls")) + glob.glob(os.path.join(self.folder_path, "*.xlsx"))
            
            if not file_paths:
                messagebox.showerror("错误", f"在选择的文件夹中没有找到Excel文件")
                return
            
            self.status_label.config(text=f"找到 {len(file_paths)} 个Excel文件，正在读取...")
            self.root.update()
            
            # 读取并合并文件
            all_data = []
            success_count = 0
            error_files = []
            
            for i, file in enumerate(file_paths):
                try:
                    df = pd.read_excel(file)
                    all_data.append(df)
                    success_count += 1
                    self.status_label.config(text=f"正在读取文件 {i+1}/{len(file_paths)}: {os.path.basename(file)}")
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
            
            # 保存结果
            output_file = os.path.join(self.folder_path, "merged_file.xlsx")
            merged_df.to_excel(output_file, index=False)
            
            # 显示成功信息
            success_message = f"合并完成！\n\n输出文件：{output_file}\n总共合并了 {len(merged_df)} 行数据\n成功读取了 {success_count} 个文件"
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