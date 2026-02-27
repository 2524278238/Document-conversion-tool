# -*- coding: utf-8 -*-
"""
PDF转图片转换器
支持将PDF文档转换为各种格式的图片
"""

import os
from pathlib import Path
import logging

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class PdfImageConverter:
    """PDF转图片转换器"""
    
    def __init__(self):
        self.check_dependencies()
        self.supported_output_formats = ['.png', '.jpg', '.jpeg', '.bmp', '.tiff']
    
    def check_dependencies(self):
        """检查依赖库是否可用"""
        self.pymupdf_available = PYMUPDF_AVAILABLE
        self.pil_available = PIL_AVAILABLE
        self.pdf2image_available = PDF2IMAGE_AVAILABLE
        
        if not any([self.pymupdf_available, self.pdf2image_available]):
            logger.warning("PDF处理库未安装，PDF转图片功能不可用")
        
        if not self.pil_available:
            logger.warning("PIL/Pillow库未安装，图片处理功能受限")
    
    def pdf_to_image(self, input_file, output_dir=None, output_format='png', dpi=150, page_range=None):
        """
        将PDF转换为图片
        
        Args:
            input_file (str): 输入的PDF文件路径
            output_dir (str): 输出目录，如果为None则使用输入文件所在目录
            output_format (str): 输出图片格式，支持png, jpg, jpeg, bmp, tiff
            dpi (int): 图片分辨率，默认150
            page_range (tuple): 页面范围，格式为(start, end)，如果为None则转换所有页面
        
        Returns:
            list: 输出文件路径列表
        
        Raises:
            Exception: 转换失败时抛出异常
        """
        if not any([self.pymupdf_available, self.pdf2image_available]):
            raise Exception("PDF转图片功能不可用，请安装PyMuPDF或pdf2image库")
        
        input_path = Path(input_file)
        if not input_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_file}")
        
        if not input_path.suffix.lower() == '.pdf':
            raise ValueError("输入文件必须是PDF文档(.pdf)")
        
        if f'.{output_format.lower()}' not in self.supported_output_formats:
            raise ValueError(f"不支持的输出格式: {output_format}")
        
        # 确定输出路径
        if output_dir is None:
            output_dir = input_path.parent
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            logger.info(f"开始转换: {input_file} -> {output_format.upper()}图片")
            
            # 优先使用PyMuPDF
            if self.pymupdf_available:
                return self._pdf_to_image_pymupdf(input_path, output_dir, output_format, dpi, page_range)
            # 备用方案：使用pdf2image
            elif self.pdf2image_available:
                return self._pdf_to_image_pdf2image(input_path, output_dir, output_format, dpi, page_range)
            else:
                raise Exception("没有可用的PDF处理库")
                
        except Exception as e:
            logger.error(f"PDF转图片失败: {str(e)}")
            raise Exception(f"PDF转图片转换失败: {str(e)}")
    
    def _pdf_to_image_pymupdf(self, input_path, output_dir, output_format, dpi, page_range):
        """
        使用PyMuPDF将PDF转换为图片
        
        Args:
            input_path (Path): 输入PDF文件路径
            output_dir (Path): 输出目录
            output_format (str): 输出格式
            dpi (int): 分辨率
            page_range (tuple): 页面范围
        
        Returns:
            list: 输出文件路径列表
        """
        output_files = []
        
        try:
            # 打开PDF文档
            pdf_document = fitz.open(str(input_path))
            total_pages = len(pdf_document)
            
            # 确定页面范围
            if page_range is None:
                start_page, end_page = 0, total_pages
            else:
                start_page, end_page = page_range
                start_page = max(0, start_page - 1)  # 转换为0基索引
                end_page = min(total_pages, end_page)
            
            logger.info(f"PDF总页数: {total_pages}, 转换页面: {start_page + 1}-{end_page}")
            
            # 计算缩放因子
            zoom = dpi / 72.0  # 72是PDF的默认DPI
            mat = fitz.Matrix(zoom, zoom)
            
            # 逐页转换
            for page_num in range(start_page, end_page):
                page = pdf_document[page_num]
                
                # 渲染页面为图片
                pix = page.get_pixmap(matrix=mat)
                
                # 生成输出文件名
                output_file = output_dir / f"{input_path.stem}_page_{page_num + 1:03d}.{output_format.lower()}"
                
                # 保存图片
                if output_format.lower() in ['jpg', 'jpeg']:
                    pix.save(str(output_file), output="jpeg")
                else:
                    pix.save(str(output_file))
                
                output_files.append(str(output_file))
                logger.info(f"已转换第 {page_num + 1} 页: {output_file}")
            
            pdf_document.close()
            
            logger.info(f"PyMuPDF转换完成，共生成 {len(output_files)} 个图片文件")
            return output_files
            
        except Exception as e:
            logger.error(f"PyMuPDF转换失败: {str(e)}")
            raise
    
    def _pdf_to_image_pdf2image(self, input_path, output_dir, output_format, dpi, page_range):
        """
        使用pdf2image将PDF转换为图片
        
        Args:
            input_path (Path): 输入PDF文件路径
            output_dir (Path): 输出目录
            output_format (str): 输出格式
            dpi (int): 分辨率
            page_range (tuple): 页面范围
        
        Returns:
            list: 输出文件路径列表
        """
        output_files = []
        
        try:
            # 确定页面范围
            first_page = None
            last_page = None
            if page_range is not None:
                first_page, last_page = page_range
            
            # 转换PDF为图片
            images = convert_from_path(
                str(input_path),
                dpi=dpi,
                first_page=first_page,
                last_page=last_page
            )
            
            logger.info(f"pdf2image转换完成，共 {len(images)} 页")
            
            # 保存图片
            for i, image in enumerate(images):
                page_num = i + 1
                if page_range is not None:
                    page_num = page_range[0] + i
                
                output_file = output_dir / f"{input_path.stem}_page_{page_num:03d}.{output_format.lower()}"
                
                # 根据格式保存
                if output_format.lower() in ['jpg', 'jpeg']:
                    image.save(str(output_file), 'JPEG', quality=95)
                else:
                    image.save(str(output_file), output_format.upper())
                
                output_files.append(str(output_file))
                logger.info(f"已保存第 {page_num} 页: {output_file}")
            
            logger.info(f"pdf2image转换完成，共生成 {len(output_files)} 个图片文件")
            return output_files
            
        except Exception as e:
            logger.error(f"pdf2image转换失败: {str(e)}")
            raise
    
    def pdf_to_single_image(self, input_file, output_file, output_format='png', dpi=150, layout='vertical'):
        """
        将PDF的所有页面合并为一张长图
        
        Args:
            input_file (str): 输入的PDF文件路径
            output_file (str): 输出图片文件路径
            output_format (str): 输出图片格式
            dpi (int): 图片分辨率
            layout (str): 布局方式，'vertical'(垂直)或'horizontal'(水平)
        
        Returns:
            str: 输出文件路径
        
        Raises:
            Exception: 转换失败时抛出异常
        """
        if not self.pil_available:
            raise Exception("图片合并功能需要Pillow库：pip install Pillow")
        
        # 先转换为单独的图片
        temp_dir = Path(input_file).parent / "temp_images"
        temp_dir.mkdir(exist_ok=True)
        
        try:
            # 转换PDF为图片
            image_files = self.pdf_to_image(input_file, str(temp_dir), output_format, dpi)
            
            if not image_files:
                raise Exception("PDF转换失败，没有生成图片")
            
            # 打开所有图片
            images = []
            for img_file in image_files:
                images.append(Image.open(img_file))
            
            # 计算合并后的图片尺寸
            if layout == 'vertical':
                total_width = max(img.width for img in images)
                total_height = sum(img.height for img in images)
            else:  # horizontal
                total_width = sum(img.width for img in images)
                total_height = max(img.height for img in images)
            
            # 创建新图片
            merged_image = Image.new('RGB', (total_width, total_height), 'white')
            
            # 粘贴图片
            current_pos = 0
            for img in images:
                if layout == 'vertical':
                    merged_image.paste(img, (0, current_pos))
                    current_pos += img.height
                else:  # horizontal
                    merged_image.paste(img, (current_pos, 0))
                    current_pos += img.width
            
            # 保存合并后的图片
            merged_image.save(output_file, output_format.upper())
            
            # 清理临时文件
            for img in images:
                img.close()
            for img_file in image_files:
                os.remove(img_file)
            temp_dir.rmdir()
            
            logger.info(f"PDF合并为单张图片完成: {output_file}")
            return output_file
            
        except Exception as e:
            # 清理临时文件
            try:
                for img_file in temp_dir.glob("*"):
                    img_file.unlink()
                temp_dir.rmdir()
            except:
                pass
            raise Exception(f"PDF合并为单张图片失败: {str(e)}")
    
    def get_supported_formats(self):
        """
        获取支持的文件格式
        
        Returns:
            dict: 支持的格式信息
        """
        available = any([self.pymupdf_available, self.pdf2image_available])
        
        return {
            'pdf_to_image': {
                'available': available,
                'input_formats': ['.pdf'],
                'output_formats': self.supported_output_formats,
                'engines': {
                    'pymupdf': self.pymupdf_available,
                    'pdf2image': self.pdf2image_available
                }
            }
        }


# 测试函数
def test_converter():
    """测试转换器功能"""
    converter = PdfImageConverter()
    
    print("支持的格式:")
    formats = converter.get_supported_formats()
    for conversion_type, info in formats.items():
        print(f"  {conversion_type}: {'可用' if info['available'] else '不可用'}")
        if info['available']:
            print(f"    输入格式: {', '.join(info['input_formats'])}")
            print(f"    输出格式: {', '.join(info['output_formats'])}")
            print(f"    可用引擎: {[k for k, v in info['engines'].items() if v]}")


if __name__ == "__main__":
    test_converter()