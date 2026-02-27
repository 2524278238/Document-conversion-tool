#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件转换工具
支持Word转PDF、PDF转Word、图片转PDF、PDF转图片等功能
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# Fix for PyInstaller "windowed" mode where stdout/stderr are None
class NullWriter:
    def write(self, text):
        pass
    def flush(self):
        pass

if sys.stdout is None:
    sys.stdout = NullWriter()
if sys.stderr is None:
    sys.stderr = NullWriter()

# 添加当前目录到Python路径
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from converters.word_pdf_converter import WordPdfConverter
from converters.image_pdf_converter import ImagePdfConverter
from converters.pdf_image_converter import PdfImageConverter
from converters.scan_converter import ScanConverter
from converters.word_md_converter import WordMdConverter


class FileConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文件转换工具")
        self.root.geometry("600x450")
        
        # 初始化转换器
        self.word_pdf_converter = WordPdfConverter()
        self.image_pdf_converter = ImagePdfConverter()
        self.pdf_image_converter = PdfImageConverter()
        self.scan_converter = ScanConverter()
        self.word_md_converter = WordMdConverter()
        
        self.setup_ui()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="文件转换工具", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 转换类型选择
        type_frame = ttk.LabelFrame(main_frame, text="选择转换类型", padding="10")
        type_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.conversion_type = tk.StringVar(value="word_to_pdf")
        
        ttk.Radiobutton(type_frame, text="Word转PDF", variable=self.conversion_type, 
                       value="word_to_pdf").grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(type_frame, text="PDF转Word", variable=self.conversion_type, 
                       value="pdf_to_word").grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(type_frame, text="图片转PDF", variable=self.conversion_type, 
                       value="image_to_pdf").grid(row=1, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(type_frame, text="PDF转图片", variable=self.conversion_type, 
                       value="pdf_to_image").grid(row=1, column=1, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(type_frame, text="图片转扫描件", variable=self.conversion_type, 
                       value="image_to_scan").grid(row=2, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(type_frame, text="Word转Markdown", variable=self.conversion_type, 
                       value="word_to_md").grid(row=2, column=1, sticky=tk.W, padx=(0, 20))
        
        # 已移除“转换选项”面板（默认行为在后台执行）
        
        # 文件选择
        file_frame = ttk.LabelFrame(main_frame, text="选择文件", padding="10")
        file_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(file_frame, text="浏览", command=self.browse_file).grid(row=0, column=1)
        
        # 输出目录选择
        output_frame = ttk.LabelFrame(main_frame, text="输出目录", padding="10")
        output_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.output_path = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_path, width=50).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(output_frame, text="浏览", command=self.browse_output).grid(row=0, column=1)
        
        # 转换按钮
        ttk.Button(main_frame, text="开始转换", command=self.convert_file).grid(row=5, column=0, columnspan=2, pady=20)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="准备就绪")
        self.status_label.grid(row=7, column=0, columnspan=2)
    
    def browse_file(self):
        conversion_type = self.conversion_type.get()
        
        if conversion_type == "word_to_pdf":
            filetypes = [("Word文档", "*.docx *.doc"), ("所有文件", "*.*")]
        elif conversion_type == "pdf_to_word":
            filetypes = [("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        elif conversion_type == "image_to_pdf":
            filetypes = [("图片文件", "*.jpg *.jpeg *.png *.bmp *.tiff"), ("所有文件", "*.*")]
        elif conversion_type == "pdf_to_image":
            filetypes = [("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        elif conversion_type == "image_to_scan":
            filetypes = [("图片文件", "*.jpg *.jpeg *.png *.bmp *.tiff"), ("所有文件", "*.*")]
        elif conversion_type == "word_to_md":
            filetypes = [("Word文档", "*.docx *.doc"), ("所有文件", "*.*")]
        else:
            filetypes = [("所有文件", "*.*")]
        
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.file_path.set(filename)
    
    def browse_output(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_path.set(directory)
    
    def convert_file(self):
        if not self.file_path.get():
            messagebox.showerror("错误", "请选择要转换的文件")
            return
        
        if not self.output_path.get():
            messagebox.showerror("错误", "请选择输出目录")
            return
        
        try:
            self.progress.start()
            self.status_label.config(text="转换中...")
            self.root.update()
            
            conversion_type = self.conversion_type.get()
            input_file = self.file_path.get()
            output_dir = self.output_path.get()
            
            if conversion_type == "word_to_pdf":
                # 使用后端默认值：accept_revisions=True, hide_comments=True
                result = self.word_pdf_converter.word_to_pdf(
                    input_file, output_dir
                )
            elif conversion_type == "pdf_to_word":
                result = self.word_pdf_converter.pdf_to_word(input_file, output_dir)
            elif conversion_type == "image_to_pdf":
                result = self.image_pdf_converter.image_to_pdf(input_file, output_dir)
            elif conversion_type == "pdf_to_image":
                result = self.pdf_image_converter.pdf_to_image(input_file, output_dir)
            elif conversion_type == "image_to_scan":
                result = self.scan_converter.convert_to_scan(input_file, output_dir)
            elif conversion_type == "word_to_md":
                result = self.word_md_converter.word_to_md(input_file, output_dir)
            
            self.progress.stop()
            self.status_label.config(text="转换完成")
            messagebox.showinfo("成功", f"文件转换完成！\n输出文件：{result}")
            
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="转换失败")
            messagebox.showerror("错误", f"转换失败：{str(e)}")


def main():
    root = tk.Tk()
    app = FileConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()