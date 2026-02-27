# -*- coding: utf-8 -*-
"""
图片扫描转换器
支持将图片转换为扫描件效果，并自动矫正视角
保留红色印章
"""

import cv2
import numpy as np
import logging
from pathlib import Path

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ScanConverter:
    """图片扫描转换器"""
    
    def __init__(self):
        self.check_dependencies()
        self.supported_formats = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']
    
    def check_dependencies(self):
        """检查依赖"""
        try:
            import cv2
            import numpy
            self.cv_available = True
        except ImportError:
            self.cv_available = False
            logger.warning("OpenCV或Numpy未安装，扫描功能不可用")

    def order_points(self, pts):
        """
        对四个点进行排序：左上，右上，右下，左下
        """
        rect = np.zeros((4, 2), dtype="float32")
        
        # 左上和右下点根据x+y的和来判断
        # 左上：和最小；右下：和最大
        s = pts.sum(axis=1)
        rect[0] = pts[np.argmin(s)]
        rect[2] = pts[np.argmax(s)]
        
        # 右上和左下点根据y-x的差来判断
        # 右上：差最小；左下：差最大
        diff = np.diff(pts, axis=1)
        rect[1] = pts[np.argmin(diff)]
        rect[3] = pts[np.argmax(diff)]
        
        return rect

    def four_point_transform(self, image, pts):
        """
        透视变换
        """
        rect = self.order_points(pts)
        (tl, tr, br, bl) = rect
        
        # 计算新图像的宽度
        widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
        widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
        maxWidth = max(int(widthA), int(widthB))
        
        # 计算新图像的高度
        heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
        heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
        maxHeight = max(int(heightA), int(heightB))
        
        # 目标点
        dst = np.array([
            [0, 0],
            [maxWidth - 1, 0],
            [maxWidth - 1, maxHeight - 1],
            [0, maxHeight - 1]], dtype="float32")
        
        # 计算变换矩阵并应用
        M = cv2.getPerspectiveTransform(rect, dst)
        warped = cv2.warpPerspective(image, M, (maxWidth, maxHeight))
        
        return warped

    def detect_document(self, image):
        """
        检测文档轮廓
        """
        # 调整大小以加快处理速度，同时保留原始比例
        ratio = image.shape[0] / 500.0
        orig = image.copy()
        
        # 调整高度为500
        height = 500
        width = int(image.shape[1] * height / image.shape[0])
        image_resized = cv2.resize(image, (width, height))
        
        # 转灰度，高斯模糊，Canny边缘检测
        gray = cv2.cvtColor(image_resized, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (5, 5), 0)
        edged = cv2.Canny(gray, 75, 200)
        
        # 寻找轮廓
        cnts = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        cnts = sorted(cnts, key=cv2.contourArea, reverse=True)[:5]
        
        screenCnt = None
        
        # 遍历轮廓
        for c in cnts:
            # 近似轮廓
            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, 0.02 * peri, True)
            
            # 如果近似轮廓有4个点，则假设找到了文档
            if len(approx) == 4:
                # 增加面积判断：如果轮廓面积小于总面积的 1/5，则认为无效
                area = cv2.contourArea(c)
                total_area = width * height
                if area < total_area * 0.2:
                    continue
                    
                screenCnt = approx
                break
        
        if screenCnt is not None:
            # 应用透视变换
            warped = self.four_point_transform(orig, screenCnt.reshape(4, 2) * ratio)
            return warped
        else:
            # 如果没找到，返回原图
            logger.warning("未检测到有效文档轮廓（面积过小或非四边形），使用原图")
            return orig

    def process_scan_effect(self, image):
        """
        处理扫描效果：增强对比度、去底色，并保留红色印章
        """
        # 1. 提取红色印章掩膜
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        
        # 红色有两个范围：0-10 和 156-180
        lower_red1 = np.array([0, 43, 46])
        upper_red1 = np.array([10, 255, 255])
        mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
        
        lower_red2 = np.array([156, 43, 46])
        upper_red2 = np.array([180, 255, 255])
        mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
        
        # 合并掩膜
        red_mask = cv2.add(mask1, mask2)
        
        # 稍微膨胀掩膜以包含印章边缘
        kernel = np.ones((3, 3), np.uint8)
        red_mask = cv2.dilate(red_mask, kernel, iterations=1)
        
        # 2. 生成扫描件效果（增强版）
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # A. 去除阴影/背景 (Background Removal)
        # 使用形态学闭操作估算背景，比高斯模糊更能保留文字结构
        # 核大小取决于图片分辨率，这里动态计算
        h, w = gray.shape
        kernel_size = int(min(h, w) * 0.02) | 1  # 约2%的宽度，确保是奇数
        if kernel_size < 3: kernel_size = 3
        
        structuring_element = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_size, kernel_size))
        bg = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, structuring_element)
        
        # 除以背景
        processed = cv2.divide(gray, bg, scale=255)
        
        # B. Gamma校正 (Gamma Correction)
        # 增强文字黑度：Gamma > 1 会让中间调变暗
        gamma = 1.5
        invGamma = 1.0 / gamma
        table = np.array([((i / 255.0) ** gamma) * 255 for i in np.arange(0, 256)]).astype("uint8")
        processed = cv2.LUT(processed, table)
        
        # C. 适度对比度增强
        # 截断极亮部分
        _, processed = cv2.threshold(processed, 230, 255, cv2.THRESH_TRUNC)
        # 归一化
        processed = cv2.normalize(processed, None, 0, 255, cv2.NORM_MINMAX)
        
        # D. 最终轻微锐化 (Unsharp Masking)
        gaussian = cv2.GaussianBlur(processed, (0, 0), 2.0)
        processed = cv2.addWeighted(processed, 1.5, gaussian, -0.5, 0)
        
        # 转回BGR以便与彩色印章合并
        scanned_bgr = cv2.cvtColor(processed, cv2.COLOR_GRAY2BGR)
        
        # 3. 合并：在红色掩膜区域使用原图，其他区域使用处理后的图
        # 将掩膜转为3通道
        mask_3ch = cv2.cvtColor(red_mask, cv2.COLOR_GRAY2BGR)
        
        # 使用np.where进行合并
        # 如果mask_3ch > 0 (是红色区域)，取原图image；否则取scanned_bgr
        result = np.where(mask_3ch > 0, image, scanned_bgr)
        
        return result

    def convert_to_scan(self, input_file, output_dir=None):
        """
        执行图片转扫描件
        """
        if not self.cv_available:
            raise Exception("OpenCV未安装，无法执行扫描转换")
            
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
            
        output_file = output_dir / f"{input_path.stem}_scan.jpg"
        
        try:
            logger.info(f"开始处理: {input_file}")
            
            # 读取图片
            # 注意：cv2.imread不支持中文路径，需要用np.fromfile读取
            img_array = np.fromfile(str(input_path), dtype=np.uint8)
            image = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
            
            if image is None:
                raise Exception("无法读取图片文件")
            
            # 1. 检测文档并矫正
            warped = self.detect_document(image)
            
            # 2. 应用扫描效果并保留红章
            result = self.process_scan_effect(warped)
            
            # 保存图片
            # cv2.imwrite不支持中文路径，使用imencode + tofile
            is_success, buffer = cv2.imencode(".jpg", result)
            if is_success:
                buffer.tofile(str(output_file))
                logger.info(f"转换成功: {output_file}")
                return str(output_file)
            else:
                raise Exception("保存图片失败")
                
        except Exception as e:
            logger.error(f"处理失败: {str(e)}")
            raise e
