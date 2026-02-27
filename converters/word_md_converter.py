# -*- coding: utf-8 -*-
"""
Word转Markdown转换器
支持.docx和.doc格式
依赖库: mammoth, markdownify, pywin32 (仅.doc转.docx需要)
"""

import os
import logging
from pathlib import Path
import shutil
import tempfile

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

try:
    import mammoth
    MAMMOTH_AVAILABLE = True
except ImportError:
    MAMMOTH_AVAILABLE = False

try:
    from markdownify import markdownify as md
    MARKDOWNIFY_AVAILABLE = True
except ImportError:
    MARKDOWNIFY_AVAILABLE = False

class WordMdConverter:
    """Word转Markdown转换器"""
    
    def __init__(self):
        self.check_dependencies()
        self.win32_available = False
        try:
            import win32com.client
            import pythoncom
            self.win32_available = True
        except ImportError:
            pass

    def check_dependencies(self):
        """检查依赖库"""
        if not MAMMOTH_AVAILABLE:
            logger.warning("mammoth库未安装，Word转Markdown功能不可用")
        if not MARKDOWNIFY_AVAILABLE:
            logger.warning("markdownify库未安装，Word转Markdown功能不可用")

    def _doc_to_docx(self, doc_path):
        """将.doc转换为临时.docx文件"""
        if not self.win32_available:
            raise ImportError("处理.doc文件需要安装pywin32库")
        
        import win32com.client
        import pythoncom
        
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        doc = None
        temp_docx = None
        try:
            doc = word.Documents.Open(str(doc_path))
            # 创建临时文件路径
            fd, temp_docx = tempfile.mkstemp(suffix='.docx')
            os.close(fd)
            # FileFormat=12 (wdFormatXMLDocument) -> .docx
            doc.SaveAs(temp_docx, FileFormat=12)
            return temp_docx
        except Exception as e:
            logger.error(f"转换.doc到.docx失败: {e}")
            raise
        finally:
            if doc:
                doc.Close()
            word.Quit()

    def word_to_md(self, input_file, output_dir=None):
        """
        将Word文档转换为Markdown
        """
        if not MAMMOTH_AVAILABLE or not MARKDOWNIFY_AVAILABLE:
            raise ImportError("请安装 mammoth 和 markdownify 库: pip install mammoth markdownify")
            
        input_path = Path(input_file).resolve()
        if not input_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_file}")
            
        # 确定输出路径
        if output_dir is None:
            output_dir = input_path.parent
        else:
            output_dir = Path(output_dir).resolve()
            output_dir.mkdir(parents=True, exist_ok=True)
            
        output_file = output_dir / f"{input_path.stem}.md"
        
        temp_file = None
        process_path = input_path
        
        try:
            # 如果是.doc格式，先转为.docx
            if input_path.suffix.lower() == '.doc':
                logger.info(f"检测到.doc格式，正在转换为临时.docx: {input_file}")
                temp_file = self._doc_to_docx(input_path)
                process_path = Path(temp_file)
            
            # 使用mammoth将docx转为html
            logger.info(f"正在读取Word内容: {process_path}")
            with open(process_path, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html = result.value
                messages = result.messages
                for message in messages:
                    logger.warning(f"Mammoth warning: {message}")
            
            # 使用markdownify将html转为markdown
            logger.info("正在转换为Markdown格式...")
            markdown_text = md(html, heading_style="ATX")
            
            # 保存文件
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(markdown_text)
                
            logger.info(f"转换成功: {output_file}")
            return str(output_file)
            
        finally:
            # 清理临时文件
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception as e:
                    logger.warning(f"清理临时文件失败: {e}")
