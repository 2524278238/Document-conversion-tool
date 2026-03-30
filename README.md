# 文件转换工具 (Document Conversion Tool)

## 🌟 为什么要做这个项目？

在日常工作和学习中，我们经常需要进行各种文件格式的转换。虽然现在市面上有很多在线的网页转换工具可以免费使用，但它们往往存在以下痛点：
1.  **文件大小限制**：免费版通常限制只能转换几兆的小文件，大文件必须充值VIP。
2.  **隐私安全风险**：将机密或私人的文档上传到未知的第三方服务器存在数据泄露风险。
3.  **随时可能收费**：很多工具前期免费“养鱼”，后期突然加上各种限制要求付费。
4.  **网络依赖**：没有网络或者网络不好的时候无法使用。

**因此，我开发了这个“点击即用”的本地文件转换工具。** 它完全免费、开源，所有转换过程都在您自己的电脑上离线完成，**没有文件大小限制，也没有隐私泄露风险**。

这是一个基于 Python 的文件格式转换工具，旨在提供简单易用的图形用户界面 (GUI)，支持多种常见的文档和图片格式转换。

> **✨ 懒人包下载**：
>
> 即使您不懂编程或没有安装 Python 环境，也可以直接下载使用！
> 本项目提供了预编译的 Windows 可执行文件 (`.exe`)，双击即可运行，无需安装任何依赖。
>
> 👉 **[下载链接](dist/神奇妙妙工具.exe)** (位于项目根目录 `dist/` 文件夹下)

## 功能特性

该工具支持以下核心功能：

*   **Word 转 PDF**: 将 Word 文档 (`.docx`, `.doc`) 转换为 PDF 文件。
*   **PDF 转 Word**: 将 PDF 文件 (`.pdf`) 转换为 Word 文档。
*   **图片 转 PDF**: 将单张或多张图片合并为一个 PDF 文件。
*   **PDF 转 图片**: 将 PDF 文件的每一页导出为图片。
*   **图片 转 扫描件**: 模拟扫描仪效果，自动矫正视角并保留红色印章。
*   **Word 转 Markdown**: 将 Word 文档 (`.docx`) 转换为 Markdown 格式。

## 快速开始

如果您不想配置 Python 环境，可以直接下载我们预编译的可执行文件。

1.  在项目根目录的 `dist/` 文件夹下找到 `神奇妙妙工具.exe`。
2.  双击运行即可使用。

## 运行环境

*   **操作系统**: Windows (推荐)
*   **Python 版本**: Python 3.8+
*   **依赖库**: 详见 `requirements.txt`

## 安装指南

1.  **克隆项目**
    ```bash
    git clone https://github.com/yourusername/file-converter.git
    cd file-converter
    ```

2.  **创建虚拟环境 (可选但推荐)**
    ```bash
    python -m venv venv
    # Windows
    venv\Scripts\activate
    ```

3.  **安装依赖**
    ```bash
    pip install -r requirements.txt
    ```

    > **注意**: 
    > *   部分功能依赖 `pywin32`，这通常需要安装 Microsoft Office Word 才能正常工作。
    > *   扫描功能依赖 `opencv-python`。
    > *   Word转Markdown功能已集成 `pypandoc-binary`，**无需手动安装 Pandoc** 即可支持公式和表格。

## 使用说明

### 启动程序

在项目根目录下运行：

```bash
python main.py
```

或者直接双击运行 `run.bat` 批处理文件。

### 操作步骤

1.  **选择转换类型**: 在主界面选择需要进行的转换操作（如“Word转PDF”）。
2.  **选择文件**: 点击“浏览”按钮选择源文件。
3.  **选择输出目录**: 指定转换后文件的保存位置。
4.  **开始转换**: 点击“开始转换”按钮，等待转换完成。

## 打包为可执行文件 (EXE)

本项目使用 `PyInstaller` 进行打包。如果需要生成独立的 `.exe` 文件，请运行：

```bash
# 使用现有的 spec 文件打包
pyinstaller FileConverter.spec
```

或者直接运行 `build_exe.bat` 脚本。

打包完成后，可执行文件将生成在 `dist/` 目录下。

## 项目结构

```
FileConverter/
├── main.py                 # 程序入口，GUI 界面
├── converters/             # 核心转换模块
│   ├── word_pdf_converter.py
│   ├── pdf_image_converter.py
│   ├── image_pdf_converter.py
│   ├── scan_converter.py   # 扫描件效果处理
│   └── word_md_converter.py
├── requirements.txt        # 项目依赖列表
├── FileConverter.spec      # PyInstaller 打包配置
└── ...
```

## 贡献指南

欢迎提交 Issue 或 Pull Request 来改进此项目。

## 许可证

[MIT License](LICENSE)
