@echo off
chcp 65001 >nul
echo ================================
echo 文件转换工具 - 依赖安装脚本
echo ================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python环境，请先安装Python 3.6或更高版本
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Python环境检查通过
echo.

echo 正在升级pip...
python -m pip install --upgrade pip

echo.
echo 正在安装基础依赖...
pip install Pillow reportlab

echo.
echo 正在安装Word转PDF依赖...
pip install docx2pdf

echo.
echo 正在安装PDF转Word依赖...
pip install pdf2docx

echo.
echo 正在安装PDF处理依赖...
pip install PyMuPDF

echo.
echo ================================
echo 安装完成！
echo ================================
echo.
echo 使用方法:
echo 1. 双击运行 run.bat 启动图形界面
echo 2. 或者在命令行中运行: python main.py
echo.
echo 如果遇到问题，请查看 README.md 文件
echo.
pause