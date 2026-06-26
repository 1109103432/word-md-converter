"""
PyInstaller 打包脚本 —— 构建单个 exe，通过命令行参数切换模式。

用法: python build.py

输出结构 (dist/):
  ├── 转换工具.exe       (~60MB, 统一入口，二合一)
  │                       双击 / 拖放 → 智能双向转换 + 剪贴板→Word
  │                       --settings → 参数配置窗口
  ├── config.json        (用户可编辑的配置文件)
  ├── 安装.bat           (终端用户安装脚本)
  └── 使用说明.txt       (简易说明)

两个桌面快捷方式指向同一个 exe：
  开始转换.lnk  → 开始转换.exe
  转换设置.lnk  → 开始转换.exe --settings

合并前: 两个 exe (116MB)，Qt 重复打包
合并后: 一个 exe (~60MB)，节省 ~56MB
"""
import subprocess
import shutil
import sys
import os
from pathlib import Path

# 确保中文输出不乱码
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

APP_VERSION = "2.2.3"

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
        "script": str(SRC / "launcher.py"),
        "name": "转换工具",
        "icon": str(ICONS / "converter.ico"),
        "desc": "Word-MD快速转换 - 拖入文件 / 剪贴板直转 / --settings 设置",
    },
]

# PyInstaller 共享参数
COMMON_ARGS = [
    "--onefile",           # 单文件输出
    "--windowed",          # 无控制台窗口
    "--noconfirm",         # 自动覆盖
    "--clean",             # 清理临时文件
    "--log-level=WARN",    # 减少输出
    # PySide6 Qt 框架
    "--hidden-import=PySide6.QtWidgets",
    "--hidden-import=PySide6.QtCore",
    "--hidden-import=PySide6.QtGui",
    # docx 相关
    "--hidden-import=docx",
    "--hidden-import=docx.opc.constants",
    # markdown
    "--hidden-import=markdown",
    # win11toast — Windows Toast 通知
    "--hidden-import=win11toast",
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
    print(f"  Word-MD快速转换 v{APP_VERSION} - PyInstaller 打包")
    print(f"  模式: 单 exe 双模式 (--settings 切换设置)")
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

    # 下载/复制 Pandoc 便携版（先放到 dist 根，稍后移入 bin/）
    _ensure_pandoc(DIST)

    # 复制 config.json 到 dist
    config_src = ROOT / "config.json"
    config_dst = DIST / "config.json"
    shutil.copy(config_src, config_dst)
    print(f"  [OK] config.json")

    # 创建 模板/ 目录（内置模板 + 用户自定义模板存放处）
    templates_dir = DIST / "模板"
    templates_dir.mkdir(exist_ok=True)

    # 复制 内置模板.docx 到 dist/模板/
    template_src = ROOT / "模板" / "内置模板.docx"
    if template_src.exists():
        shutil.copy(template_src, templates_dir / "内置模板.docx")
        print(f"  [OK] 模板/内置模板.docx")
    else:
        print(f"  [WARN] 模板/内置模板.docx 不存在 — MD→Word 将使用 Pandoc 默认样式")

    print(f"  [OK] 模板/ (自定义模板目录)")

    # ── 隐藏可执行文件到 bin/ 目录 ──
    bin_dir = DIST / "bin"
    bin_dir.mkdir(exist_ok=True)
    for exe_name in ["转换工具.exe", "pandoc.exe"]:
        src = DIST / exe_name
        if src.exists():
            shutil.move(str(src), str(bin_dir / exe_name))

    # 复制设置图标到 bin/（用于转换设置快捷方式的独立图标）
    settings_ico = ICONS / "settings.ico"
    if settings_ico.exists():
        shutil.copy(str(settings_ico), str(bin_dir / "settings.ico"))
        print(f"  [OK] bin/settings.ico")
    # Windows 隐藏 bin 目录
    try:
        import ctypes
        ctypes.windll.kernel32.SetFileAttributesW(str(bin_dir), 0x02)  # FILE_ATTRIBUTE_HIDDEN
    except Exception:
        pass
    print(f"  [OK] bin/ (隐藏的可执行文件 + 单 exe 双模式)")

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
    zip_path = ROOT / f"Word-MD快速转换-v{APP_VERSION}.zip"
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
title Word-MD快速转换 - 安装

echo.
echo ============================================
echo   Word-MD快速转换 安装程序
echo ============================================
echo.
echo   两个桌面图标共用一个程序，自动识别模式：
echo      [开始转换]    双击 = 剪贴板转Word / 拖入文件自动识别
echo      [转换设置]    双击 = 配置转换参数
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
    "$s1 = $WshShell.CreateShortcut(\"$Desktop\开始转换.lnk\"); " ^
    "$s1.TargetPath = '%APP_DIR%\bin\转换工具.exe'; " ^
    "$s1.IconLocation = '%APP_DIR%\bin\转换工具.exe,0'; " ^
    "$s1.WorkingDirectory = $Desktop; " ^
    "$s1.Description = 'Word-MD快速转换 - 拖放文件 / 双击剪贴板→Word'; " ^
    "$s1.Save(); " ^
    "" ^
    "$s2 = $WshShell.CreateShortcut(\"$Desktop\转换设置.lnk\"); " ^
    "$s2.TargetPath = '%APP_DIR%\bin\转换工具.exe'; " ^
    "$s2.Arguments = '--settings'; " ^
    "$s2.IconLocation = '%APP_DIR%\bin\settings.ico,0'; " ^
    "$s2.WorkingDirectory = $Desktop; " ^
    "$s2.Description = '配置Word-MD快速转换参数'; " ^
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
echo   桌面上已创建 2 个图标（共用一个程序）：
echo      开始转换        - 双击=剪贴板转Word / 拖入文件自动识别转换
echo      转换设置        - 配置转换参数
echo.
echo   使用方法：
echo     将 .docx 或 .md 文件拖到 [开始转换] 图标即可自动识别并转换。
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
    content = f"""Word-MD快速转换 v{APP_VERSION} — 使用说明
================================================

📌 系统要求
  - Windows 10/11 (64位)
  - 可自行安装 Microsoft Word（非必须，仅 .doc 转换需要）
  - 无需 Python 或其他依赖，解压即用

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📌 两种使用方式

  ┌──────────────────────────────────────────────┐
  │  方式一：拖入一个或多个文件（自动识别方向）              │
  │  将 .docx / .doc / .md 拖到「开始转换」图标    │
  │  .docx → .md   Word转Markdown                │
  │  .doc  → .md   旧格式（需Word），建议手动另存为docx       │
  │  .md   → .docx Markdown转Word                │
  └──────────────────────────────────────────────┘

  ┌──────────────────────────────────────────────┐
  │  方式二：双击图标（剪贴板 → Word）⭐           │
  │  复制 Markdown 内容后双击「开始转换」图标      │
  │  自动读取剪贴板 → 生成排版精致的 Word 文档     │
  │  文件名取 Markdown 首行标题，默认保存到桌面    │
  └──────────────────────────────────────────────┘

📌 桌面图标说明

  🔄 开始转换    双击=剪贴板→Word / 拖入文件=自动识别转换
  ⚙ 转换设置    双击=打开参数设置窗口
  （两个图标共用一个程序，通过命令行参数切换模式）

📌 转换结果

  • 拖入文件：转换结果保存在原文件所在目录
  • 剪贴板：结果保存到桌面（可在设置中修改）
  • 右下角弹出通知，点击通知可打开文件所在文件夹

📌 参数设置（双击「转换设置」）

  MD→Word 标签页：
  • 粘贴板导出目录 — 剪贴板→Word 的输出位置
  • 输出样式模板 — 控制 docx 的字体、颜色、行距
  • 内置通用模板 + 支持自定义模板切换
  • 标题映射 — Markdown # ⇄ Word Heading N

  Word→MD 标签页：
  • 图片提取、表格、列表、超链接 — 开关控制
  • 输出编码、添加目录占位符

  其他设置：
  • 通知开关、通知时长、是否显示输出路径

📌 关于 .doc 文件

  .doc 是 Word 97-2003 的旧格式，需本机安装 Word。
  双击图标即可转换，软件会在后台自动调用 Word
  将 .doc 转为 .docx 后再处理。
  如未安装 Word，请先另存为 .docx 再拖入。

📌 卸载

  删除本文件夹和桌面图标即可，无残留。
  所有数据仅保存在本目录，不写入系统。
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
