@echo off
chcp 65001 >nul
title Word ↔ Markdown 转换工具 - 安装

echo.
echo ============================================
echo   Word ^<--^> Markdown 转换工具 安装程序
echo ============================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.9+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [✓] Python 已安装
echo.

:: 安装依赖
echo [*] 正在安装 Python 依赖...
pip install python-docx markdown --quiet 2>&1
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络连接。
    pause
    exit /b 1
)
echo [✓] 依赖安装完成
echo.

:: 获取项目目录
set "PROJECT_DIR=%~dp0"
set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
set "SRC_DIR=%PROJECT_DIR%\src"

:: 获取 pythonw.exe 路径
for /f "tokens=*" %%i in ('python -c "import sys; print(sys.exec_prefix)"') do set "PYTHON_DIR=%%i"
set "PYTHONW=%PYTHON_DIR%\pythonw.exe"

echo [*] Python 路径：%PYTHONW%
echo [*] 项目路径：%PROJECT_DIR%
echo.

:: 获取桌面路径
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop 2^>nul') do set "DESKTOP=%%b"
if not defined DESKTOP set "DESKTOP=%USERPROFILE%\Desktop"

echo [*] 桌面路径：%DESKTOP%
echo.

:: 生成简单图标
echo [*] 正在生成图标...
python "%SRC_DIR%\generate_icons.py" "%PROJECT_DIR%" 2>nul
if errorlevel 1 (
    echo [提示] 图标生成跳过（Pillow 未安装），将使用默认图标。
)

:: 创建桌面快捷方式
echo [*] 正在创建桌面快捷方式...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$WshShell = New-Object -ComObject WScript.Shell; " ^
    "$Desktop = [Environment]::GetFolderPath('Desktop'); " ^
    "" ^
    "$Shortcut1 = $WshShell.CreateShortcut(\"$Desktop\转换工具.lnk\"); " ^
    "$Shortcut1.TargetPath = '%PYTHONW%'; " ^
    "$Shortcut1.Arguments = '\"%SRC_DIR%\converter_launcher.py\"'; " ^
    "$Shortcut1.WorkingDirectory = '%PROJECT_DIR%'; " ^
    "$Shortcut1.Description = 'Word↔Markdown 双向智能转换 - 自动识别文件类型'; " ^
    "$Shortcut1.Save(); " ^
    "" ^
    "" ^
    "$Shortcut2 = $WshShell.CreateShortcut(\"$Desktop\转换设置.lnk\"); " ^
    "$Shortcut2.TargetPath = '%PYTHONW%'; " ^
    "$Shortcut2.Arguments = '\"%SRC_DIR%\settings_app.py\"'; " ^
    "$Shortcut2.WorkingDirectory = '%PROJECT_DIR%'; " ^
    "$Shortcut2.Description = '配置Word-Markdown转换参数'; " ^
    "$Shortcut2.Save(); "

if errorlevel 1 (
    echo [错误] 快捷方式创建失败。
    pause
    exit /b 1
)
echo [✓] 桌面快捷方式已创建
echo.

:: 尝试设置图标
if exist "%PROJECT_DIR%\icons\converter.ico" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "$WshShell = New-Object -ComObject WScript.Shell; " ^
        "$Desktop = [Environment]::GetFolderPath('Desktop'); " ^
        "$Shortcut = $WshShell.CreateShortcut(\"$Desktop\转换工具.lnk\"); " ^
        "$Shortcut.IconLocation = '%PROJECT_DIR%\icons\converter.ico'; " ^
        "$Shortcut.Save(); "
)


if exist "%PROJECT_DIR%\icons\settings.ico" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "$WshShell = New-Object -ComObject WScript.Shell; " ^
        "$Desktop = [Environment]::GetFolderPath('Desktop'); " ^
        "$Shortcut = $WshShell.CreateShortcut(\"$Desktop\转换设置.lnk\"); " ^
        "$Shortcut.IconLocation = '%PROJECT_DIR%\icons\settings.ico'; " ^
        "$Shortcut.Save(); "
)

echo ============================================
echo   安装完成！
echo.
echo   桌面上已创建 2 个图标：
echo    🔄 转换工具        - 拖入 .docx 或 .md 自动识别转换
echo    ⚙ 转换设置        - 配置转换参数
echo.
echo   使用方法：将 .docx 或 .md 文件拖到 [转换工具] 图标即可自动识别并转换。
echo ============================================
echo.
pause
