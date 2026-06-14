"""
PyInstaller 打包脚本 —— 将两个入口分别构建为独立 .exe 文件。

用法: python build.py

输出结构 (dist/):
  ├── 转换工具.exe       (~20MB, 智能双向转换，自动识别文件类型)
  ├── 转换设置.exe       (~20MB, 参数配置)
  ├── config.json        (用户可编辑的配置文件)
  ├── 安装.bat           (终端用户安装脚本)
  └── 使用说明.txt       (简易说明)
"""
import subprocess
import shutil
import sys
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
BUILD = ROOT / "build"
ICONS = ROOT / "icons"
SRC = ROOT / "src"

# 清理旧构建
for d in [DIST, BUILD]:
    if d.exists():
        shutil.rmtree(d)

EXE_CONFIGS = [
    {
        "script": str(SRC / "converter_launcher.py"),
        "name": "转换工具",
        "icon": str(ICONS / "converter.ico"),
        "desc": "Word ↔ Markdown 智能转换 - 拖入文件自动识别方向",
    },
    {
        "script": str(SRC / "settings_app.py"),
        "name": "转换设置",
        "icon": str(ICONS / "settings.ico"),
        "desc": "配置Word-Markdown转换参数",
    },
]

# PyInstaller 共享参数
COMMON_ARGS = [
    "--onefile",           # 单文件输出
    "--windowed",          # 无控制台窗口
    "--noconfirm",         # 自动覆盖
    "--clean",             # 清理临时文件
    "--log-level=WARN",    # 减少输出
    # 确保 tkinter 相关模块被包含
    "--hidden-import=tkinter",
    "--hidden-import=tkinter.ttk",
    "--hidden-import=tkinter.messagebox",
    # docx 相关
    "--hidden-import=docx",
    "--hidden-import=docx.opc.constants",
    # markdown
    "--hidden-import=markdown",
    # winotify — Windows Toast 通知
    "--hidden-import=winotify",
]


def build_exe(config: dict) -> bool:
    """用 PyInstaller 构建单个 .exe。"""
    name = config["name"]
    print(f"\n{'='*60}")
    print(f"  构建: {name}")
    print(f"{'='*60}")

    args = [
        sys.executable, "-m", "PyInstaller",
        *COMMON_ARGS,
        f"--name={name}",
        f"--icon={config['icon']}",
        config["script"],
    ]

    try:
        result = subprocess.run(
            args,
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            timeout=300,
        )
        if result.returncode != 0:
            print(f"  [FAIL] PyInstaller 返回码: {result.returncode}")
            # 显示部分错误信息
            for line in result.stderr.split("\n")[-10:]:
                if line.strip():
                    print(f"    {line}")
            return False

        # 验证输出
        exe_path = DIST / f"{name}.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"  [OK]  {name}.exe  ({size_mb:.1f} MB)")
            return True
        else:
            print(f"  [FAIL] .exe 文件未找到: {exe_path}")
            return False

    except subprocess.TimeoutExpired:
        print(f"  [FAIL] 构建超时 (>5分钟)")
        return False
    except Exception as e:
        print(f"  [FAIL] {e}")
        return False


def main():
    print("=" * 60)
    print("  Word <-> Markdown 转换工具 - PyInstaller 打包")
    print("=" * 60)
    print(f"  Python: {sys.version}")
    print(f"  项目根目录: {ROOT}")
    print(f"  输出目录: {DIST}")

    # 检查依赖
    try:
        import PyInstaller  # noqa: F401
        print("  PyInstaller: 已安装")
    except ImportError:
        print("  [ERROR] PyInstaller 未安装，请运行: pip install pyinstaller")
        return 1

    # 确保图标存在
    for config in EXE_CONFIGS:
        if not Path(config["icon"]).exists():
            print(f"  [WARN] 图标不存在 — 将使用默认图标: {config['icon']}")
            config["icon"] = ""  # PyInstaller 会使用默认图标

    # 构建所有 .exe
    DIST.mkdir(parents=True, exist_ok=True)

    success = True
    for config in EXE_CONFIGS:
        if not build_exe(config):
            success = False

    if not success:
        print("\n  部分构建失败，请检查上面的错误信息。")
        return 1

    # ── 收集输出文件 ──
    print(f"\n{'='*60}")
    print("  打包分发文件")
    print(f"{'='*60}")

    # 下载/复制 Pandoc 便携版
    _ensure_pandoc(DIST)

    # 复制 config.json 到 dist
    config_src = ROOT / "config.json"
    config_dst = DIST / "config.json"
    shutil.copy(config_src, config_dst)
    print(f"  [OK] config.json")

    # 复制 template.docx 到 dist（内置参考样式模板）
    template_src = ROOT / "template.docx"
    if template_src.exists():
        shutil.copy(template_src, DIST / "template.docx")
        print(f"  [OK] template.docx")
    else:
        print(f"  [WARN] template.docx 不存在 — MD→Word 将使用 Pandoc 默认样式")

    # 创建终端用户安装脚本
    _create_installer_bat(DIST)
    print(f"  [OK] 安装.bat")

    # 创建使用说明
    _create_readme_txt(DIST)
    print(f"  [OK] 使用说明.txt")

    # ── 统计 ──
    print(f"\n{'='*60}")
    print("  构建完成")
    print(f"{'='*60}")

    total_size = sum(f.stat().st_size for f in DIST.rglob("*") if f.is_file())
    print(f"  输出目录: {DIST}")
    print(f"  总大小: {total_size / (1024*1024):.1f} MB")
    print(f"  文件列表:")
    for f in sorted(DIST.rglob("*")):
        if f.is_file():
            size = f.stat().st_size / (1024 * 1024)
            print(f"    {f.name}  ({size:.1f} MB)")

    # 创建 zip
    zip_path = ROOT / "Word-MD转换工具.zip"
    print(f"\n  正在创建压缩包...")
    shutil.make_archive(
        str(zip_path.with_suffix("")),
        "zip",
        str(DIST),
    )
    zip_size = zip_path.stat().st_size / (1024 * 1024)
    print(f"  [OK] {zip_path.name} ({zip_size:.1f} MB)")

    print(f"\n  可分发文件: {zip_path}")
    print(f"  或直接将 {DIST} 文件夹发送给用户。")
    return 0


def _create_installer_bat(dist_dir: Path):
    """生成终端用户安装脚本。"""
    content = r"""@echo off
chcp 65001 >nul
title Word ^<-^> Markdown 转换工具 - 安装

echo.
echo ============================================
echo   Word ^<-^> Markdown 转换工具 安装程序
echo ============================================
echo.
echo   本工具包含两个程序：
echo      [转换工具]    拖入 .docx 自动转 .md，拖入 .md 自动转 .docx
echo      [转换设置]    配置转换参数
echo.

:: 获取本目录
set "APP_DIR=%~dp0"
set "APP_DIR=%APP_DIR:~0,-1%"

echo   应用目录：%APP_DIR%
echo.

:: 获取桌面路径
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop 2^>nul') do set "DESKTOP=%%b"
if not defined DESKTOP set "DESKTOP=%USERPROFILE%\Desktop"

echo   桌面路径：%DESKTOP%
echo.

:: 创建快捷方式
echo   [*] 正在创建桌面快捷方式...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$WshShell = New-Object -ComObject WScript.Shell; " ^
    "$Desktop = [Environment]::GetFolderPath('Desktop'); " ^
    "" ^
    "$s1 = $WshShell.CreateShortcut(\"$Desktop\转换工具.lnk\"); " ^
    "$s1.TargetPath = '%APP_DIR%\转换工具.exe'; " ^
    "$s1.WorkingDirectory = '%APP_DIR%'; " ^
    "$s1.Description = 'Word-Markdown双向智能转换 - 拖放文件到此图标即可自动识别'; " ^
    "$s1.Save(); " ^
    "" ^
    "$s2 = $WshShell.CreateShortcut(\"$Desktop\转换设置.lnk\"); " ^
    "$s2.TargetPath = '%APP_DIR%\转换设置.exe'; " ^
    "$s2.WorkingDirectory = '%APP_DIR%'; " ^
    "$s2.Description = '配置Word-Markdown转换参数'; " ^
    "$s2.Save(); " ^
    "" ^
    "Write-Output '快捷方式创建完成'"

if errorlevel 1 (
    echo   [警告] 快捷方式创建失败，请尝试以管理员身份运行。
)

echo.
echo ============================================
echo   安装完成！
echo.
echo   桌面上已创建 2 个图标：
echo      转换工具        - 拖入 .docx 或 .md 自动识别转换
echo      转换设置        - 配置转换参数
echo.
echo   使用方法：
echo     将 .docx 或 .md 文件拖到 [转换工具] 图标即可自动识别并转换。
echo     转换后的文件保存在原文件所在目录。
echo.
echo   提示：如需卸载，删除此文件夹和桌面图标即可。
echo ============================================
echo.
pause
"""
    (dist_dir / "安装.bat").write_text(content, encoding="utf-8")


def _create_readme_txt(dist_dir: Path):
    """生成使用说明。"""
    content = """Word ↔ Markdown 转换工具 — 使用说明
========================================

📌 系统要求
  - Windows 10/11 (64位)
  - 无需安装 Python 或其他依赖

📌 使用方法

  1. 双击"安装.bat"创建桌面快捷方式

  2. 将文件拖到图标上即可自动识别转换方向：

     🔄 转换工具    ← 拖入 .docx 或 .md 文件（自动识别）
     ⚙ 转换设置     → 双击打开参数设置

  3. 转换后的文件保存在原文件所在目录
  4. 成功后右下角弹出通知，点击可打开文件夹

📌 参数设置

  双击"转换设置"可配置：
  - MD→Word：模板选择、标题映射
  - Word→MD：大纲级别检测、图片/表格/列表/超链接提取
  - 其他：通知开关、通知时长、输出路径显示

📌 配置文件

  所有设置保存在 config.json 中，可用记事本编辑。

📌 卸载

  删除本文件夹和桌面图标即可，无残留。
"""
    (dist_dir / "使用说明.txt").write_text(content, encoding="utf-8")


# Pandoc 下载配置
PANDOC_VERSION = "3.10"
PANDOC_ZIP = f"pandoc-{PANDOC_VERSION}-windows-x86_64.zip"
PANDOC_URL = f"https://github.com/jgm/pandoc/releases/download/{PANDOC_VERSION}/{PANDOC_ZIP}"


def _ensure_pandoc(dist_dir: Path):
    """确保 dist 目录中有 pandoc.exe。

    优先查找系统中已安装的 pandoc（包括便携版所在位置），
    找不到时从 GitHub 下载便携版 zip 并解压。
    """
    import urllib.request
    import zipfile

    pandoc_exe = dist_dir / "pandoc.exe"
    if pandoc_exe.exists():
        size_mb = pandoc_exe.stat().st_size / (1024 * 1024)
        print(f"  [OK] pandoc.exe ({size_mb:.1f} MB) (已有)")
        return

    # 1. 尝试从系统或上级目录复制
    for src in [
        ROOT / "pandoc.exe",
        Path("C:/Program Files/Pandoc/pandoc.exe"),
    ]:
        if src.exists():
            shutil.copy2(src, pandoc_exe)
            size_mb = pandoc_exe.stat().st_size / (1024 * 1024)
            print(f"  [OK] pandoc.exe ({size_mb:.1f} MB) (从 {src} 复制)")
            return

    # 2. 尝试从本地已下载的 zip 提取
    zip_path = ROOT / PANDOC_ZIP
    if not zip_path.exists():
        print(f"  [*] 正在下载 Pandoc {PANDOC_VERSION} ({PANDOC_URL})...")
        print(f"      下载约 40MB，请耐心等待...")
        try:
            urllib.request.urlretrieve(PANDOC_URL, str(zip_path))
            zip_size = zip_path.stat().st_size / (1024 * 1024)
            print(f"      下载完成 ({zip_size:.1f} MB)")
        except Exception as e:
            print(f"  [WARN] Pandoc 下载失败: {e}")
            print(f"    请手动下载 {PANDOC_URL}")
            print(f"    并将 pandoc.exe 放到 {dist_dir}")
            return

    # 3. 从 zip 提取 pandoc.exe
    if zip_path.exists():
        print(f"  [*] 正在解压 pandoc.exe ...")
        try:
            with zipfile.ZipFile(str(zip_path), "r") as zf:
                # pandoc zip 内的结构：pandoc-3.10/pandoc.exe
                for name in zf.namelist():
                    if name.endswith("pandoc.exe") or name == "pandoc.exe":
                        zf.extract(name, str(dist_dir))
                        # 如果有多层目录，移动到 dist 根
                        extracted = dist_dir / name
                        if extracted != pandoc_exe:
                            shutil.move(str(extracted), str(pandoc_exe))
                            # 清理空目录
                            parent = extracted.parent
                            if parent != dist_dir and parent.exists():
                                try:
                                    parent.rmdir()
                                except OSError:
                                    pass
                        break
            size_mb = pandoc_exe.stat().st_size / (1024 * 1024)
            print(f"  [OK] pandoc.exe ({size_mb:.1f} MB)")
        except Exception as e:
            print(f"  [WARN] Pandoc 解压失败: {e}")
            print(f"    MD→Word 转换将使用内置引擎（无 Pandoc）")


if __name__ == "__main__":
    sys.exit(main())
