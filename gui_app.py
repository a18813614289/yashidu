import tkinter as tk
from tkinter import ttk, filedialog
from excel_to_word import run_excel_to_word_automation
import threading

class ExcelToWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("压实度检测工具")
        self.root.geometry("600x400")
        
        self.create_widgets()
    
    def create_widgets(self):
        # 文件选择区域
        ttk.Label(self.root, text="Excel文件:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.excel_path = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.root, text="浏览...", command=self.select_excel).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(self.root, text="Word模板:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.word_path = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.word_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.root, text="浏览...", command=self.select_word).grid(row=1, column=2, padx=5, pady=5)
        
        ttk.Label(self.root, text="输出路径:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.output_path = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.output_path, width=50).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(self.root, text="浏览...", command=self.select_output).grid(row=2, column=2, padx=5, pady=5)
        
        # 参数配置区域
        ttk.Label(self.root, text="高级选项:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.root, text="数据起始行:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.start_row = tk.IntVar(value=7)
        ttk.Entry(self.root, textvariable=self.start_row, width=10).grid(row=4, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.root, text="每组行数:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.rows_per_group = tk.IntVar(value=10)
        ttk.Entry(self.root, textvariable=self.rows_per_group, width=10).grid(row=5, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.root, text="列范围:").grid(row=6, column=0, padx=5, pady=5, sticky="w")
        self.col_range = tk.StringVar(value="24-29")
        ttk.Entry(self.root, textvariable=self.col_range, width=10).grid(row=6, column=1, padx=5, pady=5, sticky="w")
        
        # 运行按钮
        self.run_button = ttk.Button(self.root, text="开始转换", command=self.run_conversion)
        self.run_button.grid(row=7, column=1, pady=20)
        
        # 状态显示
        self.status = tk.StringVar()
        self.status.set("准备就绪")
        ttk.Label(self.root, textvariable=self.status).grid(row=8, column=0, columnspan=3)
        
        # 进度条
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress.grid(row=9, column=0, columnspan=3, pady=10)
        
        # 日志区域
        ttk.Label(self.root, text="运行日志:").grid(row=10, column=0, padx=5, pady=5, sticky="w")
        self.log_text = tk.Text(self.root, height=10, width=70)
        self.log_text.grid(row=11, column=0, columnspan=3, padx=5, pady=5)
        
        # 禁用日志区域的编辑
        self.log_text.config(state=tk.DISABLED)
        
    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)
    
    def select_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word文件", "*.docx")])
        if path:
            self.word_path.set(path)
    
    def select_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word文件", "*.docx")])
        if path:
            self.output_path.set(path)
    
    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
        
    def update_progress(self, value):
        self.progress["value"] = value
        self.root.update()
    
    def validate_inputs(self):
        if not self.excel_path.get():
            raise ValueError("请选择Excel文件")
        if not self.word_path.get():
            raise ValueError("请选择Word模板文件")
        if not self.output_path.get():
            raise ValueError("请设置输出文件路径")
        
        try:
            start, end = map(int, self.col_range.get().split('-'))
            if start >= end:
                raise ValueError("列范围起始值必须小于结束值")
        except Exception:
            raise ValueError("列范围格式不正确，请使用'起始-结束'格式，如'24-29'")

    def run_conversion(self):
        def worker():
            try:
                # 验证输入
                self.validate_inputs()
                
                # 禁用运行按钮
                self.run_button.config(state=tk.DISABLED)
                self.status.set("转换中...")
                self.log_message("开始转换过程")
                
                # 重置进度条
                self.update_progress(0)
                
                # 封装原有函数调用
                def wrapped_callback(message):
                    self.log_message(message)
                    if "完成" in message:
                        self.update_progress(100)
                    elif "开始" in message:
                        self.update_progress(20)
                    elif "处理" in message:
                        self.update_progress(50)
                
                run_excel_to_word_automation(
                    self.excel_path.get(),
                    self.word_path.get(),
                    1,  # 默认复制次数
                    self.output_path.get(),
                    wrapped_callback
                )
                
                self.status.set("转换完成")
                self.log_message("转换成功完成")
                # 显示完成对话框
                tk.messagebox.showinfo("完成", "文件转换已完成！")
            except Exception as e:
                self.status.set("转换失败")
                self.log_message(f"错误: {str(e)}")
                tk.messagebox.showerror("错误", str(e))
            finally:
                # 重新启用运行按钮
                self.run_button.config(state=tk.NORMAL)
                self.update_progress(0)
        
        threading.Thread(target=worker).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToWordApp(root)
    root.mainloop()