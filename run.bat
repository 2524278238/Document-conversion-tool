@echo off
chcp 65001 >nul
echo 启动文件转换工具...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python环境
    echo 请先运行 install.bat 安装依赖
    pause
    exit /b 1
)

python main.py

if errorlevel 1 (
    echo.
    echo 程序运行出错，请检查:
    echo 1. 是否已运行 install.bat 安装依赖
    echo 2. 查看错误信息并参考 README.md
    echo.
    pause
)