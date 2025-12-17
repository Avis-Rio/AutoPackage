#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主程序 - GUI界面和主控制逻辑
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import re
from datetime import datetime
import glob

from excel_reader import AllocationTableReader
from data_transformer import DataTransformer
from template_writer import TemplateWriter
from config import FileConfig


class AutoPackageApp:
    """配分表自动转换工具主应用"""
    
    def __init__(self, root):
        """初始化应用"""
        self.root = root
        self.root.title("配分表自动转换工具 v2.1")
        self.root.geometry("900x800")
        
        # 变量定义
        self.mode_var = tk.StringVar(value="single")  # single 或 batch
        
        # 单文件模式变量
        self.allocation_file_path = tk.StringVar()
        self.template_file_path = tk.StringVar(value="template.xls")
        self.output_file_path = tk.StringVar()
        
        # 批量模式变量
        self.batch_input_dir = tk.StringVar()
        self.batch_output_dir = tk.StringVar()
        
        # 创建GUI
        self._create_widgets()
        
        # 设置默认模板路径
        default_template = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            FileConfig.DEFAULT_TEMPLATE_NAME
        )
        if os.path.exists(default_template):
            self.template_file_path.set(default_template)
            
    def _extract_id_from_filename(self, filename):
        """从文件名中提取ID（括号内的数字）"""
        basename = os.path.basename(filename)
        # 尝试匹配中文括号或英文括号内的数字
        patterns = [
            r"\((\d+)\)",  # 英文括号
            r"（(\d+)）",  # 中文括号
            r"\((\w+)\)",  # 英文括号(含字母)
            r"（(\w+)）"   # 中文括号(含字母)
        ]
        
        for pattern in patterns:
            match = re.search(pattern, basename)
            if match:
                return match.group(1)
        
        # 如果找不到，返回时间戳
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    def _generate_output_filename(self, input_path):
        """根据输入文件生成输出文件名"""
        file_id = self._extract_id_from_filename(input_path)
        return f"【箱設定 上海】{file_id} 振分.xlsx"

    def _create_widgets(self):
        """创建界面组件"""
        # 1. 标题区域
        title_frame = ttk.Frame(self.root, padding="10")
        title_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(
            title_frame,
            text="配分表自动转换工具",
            font=("Microsoft YaHei UI", 20, "bold")
        )
        title_label.pack()
        
        # 2. 模式选择区域
        mode_frame = ttk.LabelFrame(self.root, text="操作模式", padding="10")
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Radiobutton(
            mode_frame, 
            text="单文件转换", 
            variable=self.mode_var, 
            value="single",
            command=self._on_mode_change
        ).pack(side=tk.LEFT, padx=20)
        
        ttk.Radiobutton(
            mode_frame, 
            text="批量转换 (文件夹)", 
            variable=self.mode_var, 
            value="batch",
            command=self._on_mode_change
        ).pack(side=tk.LEFT, padx=20)
        
        # 3. 文件选择区域 (容器)
        self.file_container = ttk.Frame(self.root)
        self.file_container.pack(fill=tk.X, padx=10, pady=5)
        
        # 3a. 单文件选择界面
        self.single_file_frame = ttk.LabelFrame(self.file_container, text="单文件设置", padding="10")
        
        # 配分表文件
        ttk.Label(self.single_file_frame, text="配分表文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.single_file_frame, textvariable=self.allocation_file_path, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.single_file_frame, text="浏览...", command=self._browse_allocation_file).grid(row=0, column=2, padx=5, pady=5)
        
        # 模板文件 (单文件模式)
        ttk.Label(self.single_file_frame, text="模板文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.single_file_frame, textvariable=self.template_file_path, width=60).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.single_file_frame, text="浏览...", command=self._browse_template_file).grid(row=1, column=2, padx=5, pady=5)
        
        # 输出文件
        ttk.Label(self.single_file_frame, text="输出文件:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.single_file_frame, textvariable=self.output_file_path, width=60).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(self.single_file_frame, text="浏览...", command=self._browse_output_file).grid(row=2, column=2, padx=5, pady=5)
        
        # 3b. 批量选择界面
        self.batch_file_frame = ttk.LabelFrame(self.file_container, text="批量设置", padding="10")
        
        # 输入文件夹
        ttk.Label(self.batch_file_frame, text="配分表文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.batch_file_frame, textvariable=self.batch_input_dir, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.batch_file_frame, text="选择文件夹...", command=self._browse_batch_input_dir).grid(row=0, column=2, padx=5, pady=5)
        
        # 模板文件 (批量模式复用同一个变量)
        ttk.Label(self.batch_file_frame, text="模板文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.batch_file_frame, textvariable=self.template_file_path, width=60).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.batch_file_frame, text="浏览...", command=self._browse_template_file).grid(row=1, column=2, padx=5, pady=5)
        
        # 输出文件夹 (可选)
        ttk.Label(self.batch_file_frame, text="输出文件夹 (可选):").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.batch_file_frame, textvariable=self.batch_output_dir, width=60).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(self.batch_file_frame, text="选择文件夹...", command=self._browse_batch_output_dir).grid(row=2, column=2, padx=5, pady=5)
        ttk.Label(self.batch_file_frame, text="*留空则默认在配分表同级目录下创建输出文件", foreground="gray").grid(row=3, column=1, sticky=tk.W)

        # 初始化显示
        self._on_mode_change()
        
        # 4. 控制区域
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(fill=tk.X)
        
        self.convert_button = ttk.Button(
            control_frame,
            text="开始转换",
            command=self._start_process,
            style="Accent.TButton",
            width=20
        )
        self.convert_button.pack(pady=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            control_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'  # 改为确定模式以便显示具体进度
        )
        self.progress_bar.pack(fill=tk.X, padx=20, pady=5)
        
        # 状态标签
        self.status_label = ttk.Label(
            control_frame,
            text="就绪",
            font=("Microsoft YaHei UI", 10)
        )
        self.status_label.pack(pady=2)
        
        # 5. 日志输出区域
        log_frame = ttk.LabelFrame(self.root, text="处理日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=15,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 设置Tag样式
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("ERROR", foreground="red")
        self.log_text.tag_config("WARNING", foreground="orange")

    def _on_mode_change(self):
        """模式切换处理"""
        mode = self.mode_var.get()
        if mode == "single":
            self.batch_file_frame.pack_forget()
            self.single_file_frame.pack(fill=tk.X)
        else:
            self.single_file_frame.pack_forget()
            self.batch_file_frame.pack(fill=tk.X)
    
    def _browse_allocation_file(self):
        """浏览选择配分表文件"""
        filename = filedialog.askopenfilename(
            title="选择配分表文件",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.allocation_file_path.set(filename)
            # 自动设置输出文件名
            if not self.output_file_path.get():
                self._auto_set_output_path()

    def _browse_template_file(self):
        """浏览选择模板文件"""
        filename = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.template_file_path.set(filename)

    def _browse_output_file(self):
        """浏览选择输出文件"""
        filename = filedialog.asksaveasfilename(
            title="保存输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file_path.set(filename)
            
    def _browse_batch_input_dir(self):
        """选择批量输入文件夹"""
        directory = filedialog.askdirectory(title="选择配分表所在文件夹")
        if directory:
            self.batch_input_dir.set(directory)
            
    def _browse_batch_output_dir(self):
        """选择批量输出文件夹"""
        directory = filedialog.askdirectory(title="选择输出保存文件夹")
        if directory:
            self.batch_output_dir.set(directory)

    def _auto_set_output_path(self):
        """自动设置输出文件路径 (单文件模式)"""
        allocation_file = self.allocation_file_path.get()
        if allocation_file:
            dir_path = os.path.dirname(allocation_file)
            output_filename = self._generate_output_filename(allocation_file)
            output_path = os.path.join(dir_path, output_filename)
            self.output_file_path.set(output_path)

    def _start_conversion(self):
        """开始单文件转换"""
        # 验证输入
        if not self.allocation_file_path.get():
            messagebox.showerror("错误", "请选择配分表文件")
            return
        
        if not self.output_file_path.get():
            messagebox.showerror("错误", "请选择输出文件路径")
            return
            
        # 禁用转换按钮
        self.convert_button.config(state=tk.DISABLED)
        
        # 清空日志
        self.log_text.delete("1.0", tk.END)
        
        # 启动进度条
        self.progress_bar.config(mode='indeterminate')
        self.progress_bar.start(10)
        
        # 在新线程中执行转换
        thread = threading.Thread(target=self._do_conversion)
        thread.daemon = True
        thread.start()

    def _start_process(self):
        """开始处理流程"""
        mode = self.mode_var.get()
        
        # 验证模板文件
        if not self.template_file_path.get():
            messagebox.showerror("错误", "请选择模板文件")
            return
            
        if mode == "single":
            self._start_conversion()
        else:
            self._start_batch_conversion()

    def _start_batch_conversion(self):
        """开始批量转换"""
        input_dir = self.batch_input_dir.get()
        if not input_dir or not os.path.exists(input_dir):
            messagebox.showerror("错误", "请选择有效的配分表文件夹")
            return
            
        # 获取所有excel文件
        files = []
        for ext in ['*.xls', '*.xlsx']:
            files.extend(glob.glob(os.path.join(input_dir, ext)))
        
        # 过滤临时文件
        files = [f for f in files if not os.path.basename(f).startswith('~$')]
        
        if not files:
            messagebox.showwarning("警告", "该文件夹下没有找到Excel文件")
            return
            
        # 禁用按钮，重置界面
        self.convert_button.config(state=tk.DISABLED)
        self.log_text.delete("1.0", tk.END)
        self.progress_var.set(0)
        self.progress_bar.config(mode='determinate', maximum=len(files))
        
        # 启动线程
        thread = threading.Thread(target=self._do_batch_conversion, args=(files,))
        thread.daemon = True
        thread.start()

    def _ask_overwrite_action(self, filename):
        """
        询问覆盖操作
        返回: 'skip' (跳过/保留原有), 'overwrite' (覆盖)
        """
        # 使用StringVar在主线程和子线程间传递结果
        result_var = tk.StringVar(value="")
        
        def _ask():
            res = messagebox.askyesno(
                "文件已存在",
                f"输出文件已存在：\n{filename}\n\n是否覆盖？\n\n点击'是'覆盖，点击'否'跳过（保留原文件）。"
            )
            result_var.set("overwrite" if res else "skip")
        
        # 在主线程中调用messagebox
        self.root.after(0, _ask)
        
        # 等待用户响应
        self.root.wait_variable(result_var)
        return result_var.get()

    def _do_batch_conversion(self, files):
        """执行批量转换"""
        try:
            total_files = len(files)
            success_count = 0
            fail_count = 0
            skipped_count = 0
            
            self._log(f"开始批量处理，共找到 {total_files} 个文件", "INFO")
            self._log("=" * 60)
            
            output_base_dir = self.batch_output_dir.get()
            
            for index, file_path in enumerate(files):
                filename = os.path.basename(file_path)
                self._update_status(f"正在处理 ({index+1}/{total_files}): {filename}")
                self._log(f"正在处理: {filename}")
                
                try:
                    # 确定输出路径
                    if output_base_dir:
                        output_dir = output_base_dir
                    else:
                        output_dir = os.path.dirname(file_path)
                        
                    output_filename = self._generate_output_filename(file_path)
                    output_path = os.path.join(output_dir, output_filename)
                    
                    # 检查文件是否存在
                    if os.path.exists(output_path):
                        action = self._ask_overwrite_action(output_filename)
                        if action == 'skip':
                            self._log(f"  -> 跳过: 输出文件已存在且用户选择保留", "WARNING")
                            skipped_count += 1
                            self.progress_var.set(index + 1)
                            continue
                    
                    # 执行转换
                    self._perform_single_conversion(file_path, self.template_file_path.get(), output_path)
                    
                    self._log(f"  -> 成功生成: {output_filename}", "SUCCESS")
                    success_count += 1
                    
                except Exception as e:
                    self._log(f"  -> 处理失败: {str(e)}", "ERROR")
                    fail_count += 1
                    import traceback
                    print(traceback.format_exc())
                
                # 更新进度
                self.progress_var.set(index + 1)
                
            self._log("=" * 60)
            summary = f"批量处理完成！成功: {success_count}, 失败: {fail_count}, 跳过: {skipped_count}"
            self._log(summary, "INFO" if fail_count == 0 else "WARNING")
            self._update_status("处理完成")
            
            self.root.after(0, lambda: messagebox.showinfo("完成", summary))
            
        except Exception as e:
            self._log(f"批量处理发生致命错误: {str(e)}", "ERROR")
            
        finally:
            self.root.after(0, lambda: self.convert_button.config(state=tk.NORMAL))

    def _perform_single_conversion(self, input_path, template_path, output_path):
        """执行单个文件转换逻辑 (复用逻辑)"""
        self._log(f"读取配分表: {os.path.basename(input_path)}")
        
        # 1. 读取配分表
        reader = AllocationTableReader(input_path)
        allocation_data = reader.read()
        reader.close()
        
        products_count = len(allocation_data['products'])
        self._log(f"  -> 成功读取 {products_count} 个品番的数据")
        
        # 2. 数据转换
        self._log("  -> 正在进行数据转换和PT分组...")
        transformer = DataTransformer(allocation_data)
        transformed_data = transformer.transform()
        
        sku_count = len(transformed_data['skus'])
        pt_count = len(transformed_data['pt_groups'])
        self._log(f"  -> 分析完成: {sku_count} 个SKU, {pt_count} 个PT分组")
        
        # 3. 写入输出文件
        self._log(f"  -> 正在生成Excel文件...")
        writer = TemplateWriter(template_path, output_path)
        writer.write(transformed_data)
        self._log(f"  -> 文件写入完成")
        
    def _do_conversion(self):
        """执行单文件转换（在后台线程中）"""
        try:
            self.progress_bar.config(mode='indeterminate')
            self.progress_bar.start(10)
            self._update_status("正在处理...")
            
            input_path = self.allocation_file_path.get()
            template_path = self.template_file_path.get()
            output_path = self.output_file_path.get()
            
            self._log(f"开始处理单个文件: {os.path.basename(input_path)}")
            
            self._perform_single_conversion(input_path, template_path, output_path)
            
            self._log(f"转换成功！输出文件: {output_path}", "SUCCESS")
            self._update_status("转换完成")
            
            self.root.after(0, lambda: messagebox.showinfo("成功", f"转换完成！\n\n输出文件:\n{output_path}"))
            
        except Exception as e:
            error_msg = f"转换失败: {str(e)}"
            self._log(error_msg, "ERROR")
            self._update_status("转换失败")
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            
        finally:
            self.root.after(0, self.progress_bar.stop)
            self.root.after(0, lambda: self.convert_button.config(state=tk.NORMAL))

    def _update_status(self, status: str):
        """更新状态标签"""
        self.root.after(0, lambda: self.status_label.config(text=status))

    def _log(self, message: str, level: str = "INFO"):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        def _append():
            self.log_text.insert(tk.END, log_message, level)
            self.log_text.see(tk.END)
            
        self.root.after(0, _append)


def main():
    """主函数"""
    root = tk.Tk()
    app = AutoPackageApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
