# -*- coding: utf-8 -*-
"""
Word与PDF互转转换器
支持Word转PDF、PDF转Word
"""

import os
from pathlib import Path
import logging
import traceback
import win32com.client
import pythoncom

try:
    from docx2pdf import convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

try:
    from pdf2docx import Converter
    PDF2DOCX_AVAILABLE = True
except ImportError:
    PDF2DOCX_AVAILABLE = False

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class WordPdfConverter:
    """Word与PDF互转转换器"""
    
    def __init__(self):
        self.check_dependencies()
    
    def check_dependencies(self):
        """检查依赖库是否可用"""
        self.docx2pdf_available = DOCX2PDF_AVAILABLE
        self.pdf2docx_available = PDF2DOCX_AVAILABLE
        
        # Word转PDF现在主要依赖pywin32，docx2pdf作为备选
        try:
            import win32com.client
            self.win32_available = True
        except ImportError:
            self.win32_available = False
            logger.warning("pywin32库未安装，Word转PDF功能可能受限")
        
        if not self.pdf2docx_available:
            logger.warning("pdf2docx库未安装，PDF转Word功能不可用")
    
    def word_to_pdf(self, input_file, output_dir=None):
        """
        将Word文档转换为PDF
        
        Args:
            input_file (str): 输入的Word文件路径
            output_dir (str): 输出目录，如果为None则使用输入文件所在目录
        
        Returns:
            str: 输出文件路径
        """
        if not self.win32_available and not self.docx2pdf_available:
            raise Exception("Word转PDF功能不可用，请安装pywin32或docx2pdf库")
        
        input_path = Path(input_file).resolve()
        if not input_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_file}")
        
        if not input_path.suffix.lower() in ['.docx', '.doc']:
            raise ValueError(f"不支持的文件格式: {input_path.suffix}")
        
        # 确定输出路径
        if output_dir is None:
            output_dir = input_path.parent
        else:
            output_dir = Path(output_dir).resolve()
            output_dir.mkdir(parents=True, exist_ok=True)
        
        output_file = output_dir / f"{input_path.stem}.pdf"
        
        word = None
        doc = None
        try:
            logger.info(f"开始转换: {input_file} -> {output_file}")
            
            # 初始化COM环境（多线程/PyInstaller环境中必须）
            pythoncom.CoInitialize()
            
            # 使用win32com直接调用Word，比docx2pdf更可控且支持.doc
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            # 打开文档
            doc = word.Documents.Open(str(input_path))
            
            # 保存为PDF (FileFormat=17)
            doc.SaveAs(str(output_file), FileFormat=17)
            
            # 关闭文档
            doc.Close()
            doc = None
            
            # 退出Word
            word.Quit()
            word = None
            
            if output_file.exists():
                logger.info(f"转换成功: {output_file}")
                return str(output_file)
            else:
                raise Exception("转换失败，输出文件未生成")
                
        except Exception as e:
            # 记录详细堆栈信息
            error_msg = traceback.format_exc()
            logger.error(f"转换失败: {str(e)}\n{error_msg}")
            
            # 尝试清理资源
            try:
                if doc: doc.Close(SaveChanges=False)
                if word: word.Quit()
            except:
                pass
                
            raise Exception(f"转换失败: {str(e)}")
        finally:
            # 释放COM资源
            pythoncom.CoUninitialize()

    def pdf_to_word(self, input_file, output_dir=None):
        """
        将PDF文档转换为Word
        
        Args:
            input_file (str): 输入的PDF文件路径
            output_dir (str): 输出目录，如果为None则使用输入文件所在目录
        
        Returns:
            str: 输出文件路径
        """
        if not self.pdf2docx_available:
            raise Exception("PDF转Word功能不可用，请安装pdf2docx库：pip install pdf2docx")
        
        input_path = Path(input_file)
        if not input_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_file}")
        
        if not input_path.suffix.lower() == '.pdf':
            raise ValueError("输入文件必须是PDF文档(.pdf)")
        
        # 确定输出路径
        if output_dir is None:
            output_dir = input_path.parent
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        
        output_file = output_dir / f"{input_path.stem}.docx"
        
        try:
            logger.info(f"开始转换: {input_file} -> {output_file}")
            
            cv = Converter(str(input_path))
            cv.convert(str(output_file), start=0, end=None)
            cv.close()
            
            if output_file.exists():
                logger.info(f"转换成功: {output_file}")
                return str(output_file)
            else:
                raise Exception("转换失败，输出文件未生成")
                
        except Exception as e:
            logger.error(f"转换失败: {str(e)}")
            raise e
