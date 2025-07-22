@echo off
chcp 65001 >nul
title BIRT AI Excel转换器

echo.
echo ==========================================
echo      🚀 BIRT AI Excel转换器
echo ==========================================
echo.

if "%~1"=="" (
    echo 📖 使用方法:
    echo 1. 将Excel文件拖拽到此批处理文件上
    echo 2. 或者直接运行此文件进入交互模式
    echo.
    echo 💡 提示: 支持 .xlsx 和 .xls 格式
    echo.
    python upload_excel.py --interactive
) else (
    echo 📁 检测到拖拽文件，开始处理...
    echo.
    python upload_excel.py %*
)

echo.
echo 📋 处理完成! 按任意键查看结果目录...
pause >nul

if exist "output" (
    explorer output
) else (
    echo ⚠️ 输出目录不存在，可能处理过程中出现错误
)

echo.
echo 🎉 感谢使用! 按任意键退出...
pause >nul