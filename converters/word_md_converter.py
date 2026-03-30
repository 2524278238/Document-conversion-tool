# -*- coding: utf-8 -*-
"""
Word转Markdown转换器
支持.docx和.doc格式
依赖库: mammoth, markdownify, pywin32 (仅.doc转.docx需要), pypandoc (推荐，用于更好支持表格和公式)
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

try:
    import pypandoc
    PYPANDOC_AVAILABLE = True
except ImportError:
    PYPANDOC_AVAILABLE = False

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
            logger.warning("mammoth库未安装，作为备用转换方案不可用")
        if not MARKDOWNIFY_AVAILABLE:
            logger.warning("markdownify库未安装，作为备用转换方案不可用")
        
        self.pandoc_available = False
        if PYPANDOC_AVAILABLE:
            try:
                # 尝试获取 pandoc 版本以确认是否安装了 pandoc 可执行文件
                # 可能会抛出 OSError 或其他异常
                version = pypandoc.get_pandoc_version()
                self.pandoc_available = True
                logger.info(f"检测到 Pandoc (版本 {version})，将优先使用 Pandoc 进行转换（支持表格和公式）")
            except OSError:
                logger.info("未在系统路径中检测到 Pandoc，尝试自动下载...")
                try:
                    pypandoc.download_pandoc()
                    version = pypandoc.get_pandoc_version()
                    self.pandoc_available = True
                    logger.info(f"Pandoc 下载并安装成功 (版本 {version})")
                except Exception as e:
                     logger.warning(f"自动下载 Pandoc 失败: {e}。将回退到使用 mammoth (公式支持较弱)")
                     logger.warning("请手动安装 Pandoc 以获得最佳转换效果: https://pandoc.org/installing.html")
            except Exception as e:
                logger.warning(f"检测到 pypandoc 库，但无法调用 pandoc 可执行文件: {e}。将回退到使用 mammoth (公式支持较弱)")
                logger.warning("请安装 Pandoc 以获得最佳转换效果: https://pandoc.org/installing.html")
        else:
            logger.info("未检测到 pypandoc 库，将使用 mammoth 进行转换 (公式支持较弱)")

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
        if not (MAMMOTH_AVAILABLE and MARKDOWNIFY_AVAILABLE) and not self.pandoc_available:
             raise ImportError("未找到可用的转换库。请安装 pypandoc (推荐) 或 mammoth + markdownify。")
            
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
            else:
                process_path = input_path

            # 优先尝试使用 Pandoc
            if self.pandoc_available:
                try:
                    return self._word_to_md_pandoc(process_path, output_file)
                except Exception as e:
                    logger.error(f"Pandoc 转换失败: {e}，尝试使用备用方案 (Mammoth)")
            
            # 备用方案：Mammoth + Markdownify
            return self._word_to_md_mammoth(process_path, output_file)
            
        finally:
            # 清理临时文件
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception as e:
                    logger.warning(f"清理临时文件失败: {e}")

    def _word_to_md_pandoc(self, input_path, output_file):
        """使用 Pandoc 转换"""
        logger.info(f"正在使用 Pandoc 转换: {input_path}")
        
        # 提取图片到同级目录的 _media 文件夹
        media_dir = output_file.parent / f"{input_path.stem}_media"
        
        # pypandoc.convert_file 返回空字符串表示成功（当指定outputfile时）
        pypandoc.convert_file(
            str(input_path), 
            'markdown', 
            outputfile=str(output_file),
            extra_args=['--wrap=none', f'--extract-media={str(media_dir)}']
        )
        logger.info(f"Pandoc 转换成功: {output_file}")
        return str(output_file)

    def _word_to_md_mammoth(self, input_path, output_file):
        """使用 Mammoth 转换"""
        if not MAMMOTH_AVAILABLE or not MARKDOWNIFY_AVAILABLE:
            raise ImportError("请安装 mammoth 和 markdownify 库: pip install mammoth markdownify")
            
        # 使用mammoth将docx转为html
        logger.info(f"正在读取Word内容 (Mammoth): {input_path}")
        with open(input_path, "rb") as docx_file:
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
