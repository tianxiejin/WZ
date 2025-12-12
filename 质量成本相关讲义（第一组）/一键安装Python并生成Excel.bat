@echo off
chcp 65001 >nul
echo ========================================
echo Python 自动安装脚本
echo ========================================
echo.

REM 检查是否有管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 错误：需要管理员权限运行此脚本
    echo 请右键点击此文件，选择"以管理员身份运行"
    pause
    exit /b 1
)

echo 正在检查Python是否已安装...
python --version >nul 2>&1
if %errorLevel% equ 0 (
    echo Python已安装！版本：
    python --version
    echo.
    goto INSTALL_OPENPYXL
)

echo Python未安装，开始安装...
echo.
echo 尝试使用winget安装Python...
winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements

if %errorLevel% neq 0 (
    echo.
    echo winget安装失败。
    echo.
    echo 请手动安装Python：
    echo 1. 访问 https://www.python.org/downloads/
    echo 2. 下载并运行安装程序
    echo 3. 务必勾选 "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo.
echo Python安装完成！
echo 请关闭此窗口，重新打开命令提示符或VSCode终端
echo 然后再次运行此脚本
pause
exit /b 0

:INSTALL_OPENPYXL
echo.
echo ========================================
echo 步骤2：安装openpyxl库
echo ========================================
echo.
python -m pip install --upgrade pip
python -m pip install openpyxl

if %errorLevel% neq 0 (
    echo.
    echo openpyxl安装失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 步骤3：运行Excel生成脚本
echo ========================================
echo.
cd /d "%~dp0"
python "生成ABC成本模型Excel.py"

if %errorLevel% equ 0 (
    echo.
    echo ========================================
    echo 成功！Excel模型已生成
    echo ========================================
    echo.
    echo 文件位置：瓦轴集团ABC成本模型_演示版.xlsx
    echo.
) else (
    echo.
    echo 脚本运行失败，请检查错误信息
)

echo.
pause
