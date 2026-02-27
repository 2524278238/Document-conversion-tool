# -*- coding: utf-8 -*-
"""
图片转PDF转换器
支持将各种格式的图片转换为PDF
"""

import os
from pathlib import Path
import logging

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.utils import ImageReader
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ImagePdfConverter:
    """图片转PDF转换器"""
    
    def __init__(self):
        self.check_dependencies()
        self.supported_formats = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.gif']
    
    def check_dependencies(self):
        """检查依赖库是否可用"""
        self.pil_available = PIL_AVAILABLE
        self.reportlab_available = REPORTLAB_AVAILABLE
        
        if not self.pil_available:
            logger.warning("PIL/Pillow库未安装，图片处理功能不可用")
        
        if not self.reportlab_available:
            logger.warning("ReportLab库未安装，PDF生成功能不可用")
    
    def image_to_pdf(self, input_file, output_dir=None, page_size='A4'):
        """
        将图片转换为PDF
        
        Args:
            input_file (str): 输入的图片文件路径
            output_dir (str): 输出目录，如果为None则使用输入文件所在目录
            page_size (str): 页面大小，支持'A4', 'letter'等
        
        Returns:
            str: 输出文件路径
        
        Raises:
            Exception: 转换失败时抛出异常
        """
        if not self.pil_available:
            raise Exception("图片转PDF功能不可用，请安装Pillow库：pip install Pillow")
        
        input_path = Path(input_file)
        if not input_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_file}")
        
        if not input_path.suffix.lower() in self.supported_formats:
            raise ValueError(f"不支持的图片格式: {input_path.suffix}")
        
        # 确定输出路径
        if output_dir is None:
            output_dir = input_path.parent
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        
        output_file = output_dir / f"{input_path.stem}.pdf"
        
        try:
            logger.info(f"开始转换: {input_file} -> {output_file}")
            
            # 使用PIL进行转换（简单方法）
            with Image.open(input_path) as img:
                # 转换为RGB模式（PDF需要）
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # 保存为PDF
                img.save(str(output_file), "PDF", resolution=100.0)
            
            if output_file.exists():
                logger.info(f"转换成功: {output_file}")
                return str(output_file)
            else:
                raise Exception("转换失败，输出文件未生成")
                
        except Exception as e:
            logger.error(f"图片转PDF失败: {str(e)}")
            # 尝试使用ReportLab作为备用方案
            if self.reportlab_available:
                return self._image_to_pdf_reportlab(input_file, output_dir, page_size)
            else:
                raise Exception(f"图片转PDF转换失败: {str(e)}")
    
    def _image_to_pdf_reportlab(self, input_file, output_dir, page_size='A4'):
        """
        使用ReportLab将图片转换为PDF（备用方案）
        
        Args:
            input_file (str): 输入的图片文件路径
            output_dir (str): 输出目录
            page_size (str): 页面大小
        
        Returns:
            str: 输出文件路径
        """
        input_path = Path(input_file)
        output_file = Path(output_dir) / f"{input_path.stem}.pdf"
        
        try:
            logger.info(f"使用ReportLab转换: {input_file} -> {output_file}")
            
            # 获取页面大小
            if page_size.upper() == 'A4':
                page_width, page_height = A4
            elif page_size.upper() == 'LETTER':
                page_width, page_height = letter
            else:
                page_width, page_height = A4
            
            # 创建PDF
            c = canvas.Canvas(str(output_file), pagesize=(page_width, page_height))
            
            # 打开图片获取尺寸
            with Image.open(input_path) as img:
                img_width, img_height = img.size
                
                # 计算缩放比例以适应页面
                scale_x = page_width / img_width
                scale_y = page_height / img_height
                scale = min(scale_x, scale_y) * 0.9  # 留一些边距
                
                # 计算居中位置
                new_width = img_width * scale
                new_height = img_height * scale
                x = (page_width - new_width) / 2
                y = (page_height - new_height) / 2
                
                # 添加图片到PDF
                c.drawImage(str(input_path), x, y, width=new_width, height=new_height)
            
            c.save()
            
            if output_file.exists():
                logger.info(f"ReportLab转换成功: {output_file}")
                return str(output_file)
            else:
                raise Exception("ReportLab转换失败，输出文件未生成")
                
        except Exception as e:
            logger.error(f"ReportLab转换失败: {str(e)}")
            raise Exception(f"图片转PDF转换失败: {str(e)}")
    
    def images_to_pdf(self, input_files, output_file, page_size='A4'):
        """
        将多个图片合并为一个PDF
        
        Args:
            input_files (list): 输入的图片文件路径列表
            output_file (str): 输出PDF文件路径
            page_size (str): 页面大小
        
        Returns:
            str: 输出文件路径
        
        Raises:
            Exception: 转换失败时抛出异常
        """
        if not self.pil_available:
            raise Exception("图片转PDF功能不可用，请安装Pillow库：pip install Pillow")
        
        if not input_files:
            raise ValueError("输入文件列表不能为空")
        
        # 检查所有输入文件
        for input_file in input_files:
            input_path = Path(input_file)
            if not input_path.exists():
                raise FileNotFoundError(f"输入文件不存在: {input_file}")
            if not input_path.suffix.lower() in self.supported_formats:
                raise ValueError(f"不支持的图片格式: {input_path.suffix}")
        
        try:
            logger.info(f"开始合并 {len(input_files)} 个图片到PDF: {output_file}")
            
            images = []
            for input_file in input_files:
                with Image.open(input_file) as img:
                    # 转换为RGB模式
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    images.append(img.copy())
            
            # 保存为PDF
            if images:
                images[0].save(
                    output_file, 
                    "PDF", 
                    resolution=100.0, 
                    save_all=True, 
                    append_images=images[1:]
                )
            
            output_path = Path(output_file)
            if output_path.exists():
                logger.info(f"合并成功: {output_file}")
                return str(output_file)
            else:
                raise Exception("合并失败，输出文件未生成")
                
        except Exception as e:
            logger.error(f"图片合并PDF失败: {str(e)}")
            raise Exception(f"图片合并PDF转换失败: {str(e)}")
    
    def get_supported_formats(self):
        """
        获取支持的文件格式
        
        Returns:
            dict: 支持的格式信息
        """
        return {
            'image_to_pdf': {
                'available': self.pil_available,
                'input_formats': self.supported_formats,
                'output_format': '.pdf'
            }
        }


# 测试函数
def test_converter():
    """测试转换器功能"""
    converter = ImagePdfConverter()
    
    print("支持的格式:")
    formats = converter.get_supported_formats()
    for conversion_type, info in formats.items():
        print(f"  {conversion_type}: {'可用' if info['available'] else '不可用'}")
        if info['available']:
            print(f"    输入格式: {', '.join(info['input_formats'])}")
            print(f"    输出格式: {info['output_format']}")


if __name__ == "__main__":
    test_converter()