@echo off
chcp 65001 >nul
title 快速修复依赖 - 电缆设计软件

echo.
echo ========================================
echo   快速修复依赖 - 电缆设计软件
echo ========================================
echo.

REM 检查Python环境
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 错误：未检测到Python环境
    echo.
    echo 请先安装Python 3.6或更高版本
    echo 下载地址：https://www.python.org/downloads/
    echo.
    timeout /t 5 >nul
    exit /b 1
)

echo 正在检查和安装依赖包...
echo.

REM 升级pip
echo [1/4] 升级pip...
python -m pip install --upgrade pip

REM 安装pandas
echo [2/4] 安装pandas...
python -m pip install pandas
if %errorlevel% neq 0 (
    echo ❌ pandas 安装失败
    goto :error
)

REM 安装openpyxl
echo [3/4] 安装openpyxl...
python -m pip install openpyxl
if %errorlevel% neq 0 (
    echo ❌ openpyxl 安装失败
    goto :error
)

REM 验证安装
echo [4/4] 验证安装...
python -c "import pandas, openpyxl; print('✓ 所有依赖包安装成功')"
if %errorlevel% neq 0 (
    echo ❌ 依赖验证失败
    goto :error
)

echo.
echo ✅ 依赖修复完成！
echo.
echo 现在可以正常启动电缆设计软件了
echo.

REM 询问是否启动软件
set /p choice="是否立即启动软件？(y/n): "
if /i "%choice%"=="y" (
    echo.
    echo 启动软件...
    python cable_design_system_v4.py
) else if /i "%choice%"=="yes" (
    echo.
    echo 启动软件...
    python cable_design_system_v4.py
)

exit /b 0

:error
echo.
echo ❌ 依赖安装失败
echo.
echo 请尝试以下解决方案：
echo 1. 以管理员身份运行此脚本
echo 2. 检查网络连接
echo 3. 手动运行：python install_dependencies.py
echo 4. 联系技术支持
echo.
timeout /t 10 >nul
exit /b 1