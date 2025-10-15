import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from main_module import run_excel_to_word_automation

class App:
    def __init__(self, master):
        self.master = master
        master.title("压实度检测报告生成工具")
        
        # 添加标题标签
        title_frame = ttk.Frame(master)
        title_frame.pack(fill="x", padx=5, pady=5)
        
        self.title_label = ttk.Label(
            title_frame,
            text="压实度检测报告生成工具",
            font=("Microsoft YaHei", 14, "bold"),
            foreground="black",
            background="#f0f0f0"  # 浅灰色背景
        )
        self.title_label.pack(fill="x", pady=10)
        
        # 主内容框架
        self.main_frame = ttk.Frame(master)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        master.geometry("600x400")
        
        # 输入文件选择
        excel_frame = ttk.Frame(self.main_frame)
        excel_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(excel_frame, text="Excel文件路径:").pack(side="left", padx=5)
        self.excel_path = tk.StringVar()
        ttk.Entry(excel_frame, textvariable=self.excel_path, width=50).pack(side="left", expand=True, fill="x", padx=5)
        ttk.Button(excel_frame, text="浏览", command=self.select_excel).pack(side="left", padx=5)
        
        # Word模板选择
        word_frame = ttk.Frame(self.main_frame)
        word_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(word_frame, text="Word模板路径:").pack(side="left", padx=5)
        self.word_path = tk.StringVar()
        ttk.Entry(word_frame, textvariable=self.word_path, width=50).pack(side="left", expand=True, fill="x", padx=5)
        ttk.Button(word_frame, text="浏览", command=self.select_word).pack(side="left", padx=5)
        
        # 输出路径选择
        output_frame = ttk.Frame(self.main_frame)
        output_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(output_frame, text="输出文件路径:").pack(side="left", padx=5)
        self.output_path = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_path, width=50).pack(side="left", expand=True, fill="x", padx=5)
        ttk.Button(output_frame, text="浏览", command=self.select_output).pack(side="left", padx=5)
        
        # 状态显示
        self.status = tk.Text(self.main_frame, height=10)
        self.status.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 执行按钮
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill="x", pady=10)
        ttk.Button(button_frame, text="生成报告", command=self.run_conversion).pack()
        
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
            
    def log_status(self, message):
        self.status.insert(tk.END, message + "\n")
        self.status.see(tk.END)
        self.master.update()
        
    def run_conversion(self):
        try:
            if not all([self.excel_path.get(), self.word_path.get(), self.output_path.get()]):
                messagebox.showerror("错误", "请填写所有文件路径")
                return
                
            self.log_status("开始生成报告...")
            run_excel_to_word_automation(
                self.excel_path.get(),
                self.word_path.get(),
                1,  # copy_count
                self.output_path.get(),
                self.log_status
            )
            self.log_status("报告生成完成！")
            messagebox.showinfo("成功", "报告生成完成")
        except Exception as e:
            self.log_status(f"错误: {str(e)}")
            messagebox.showerror("错误", f"生成报告时出错: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()