@echo off
echo Installing requirements...
pip install -r requirements.txt

echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

echo Building EXE...
python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "FileConverter" ^
    --icon=NONE ^
    --hidden-import "converters" ^
    --hidden-import "converters.word_pdf_converter" ^
    --hidden-import "converters.image_pdf_converter" ^
    --hidden-import "converters.pdf_image_converter" ^
    --hidden-import "converters.scan_converter" ^
    --hidden-import "docx2pdf" ^
    --hidden-import "pdf2docx" ^
    --hidden-import "cv2" ^
    --hidden-import "PIL" ^
    --hidden-import "fitz" ^
    main.py

echo.
if exist "dist\FileConverter.exe" (
    echo Build successful!
    echo The executable file is located at: dist\FileConverter.exe
) else (
    echo Build failed.
)
pause
