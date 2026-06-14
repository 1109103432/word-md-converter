"""创建桌面快捷方式（Python 实现，避免 PowerShell 编码问题）。"""
import os
import sys
from pathlib import Path

DESKTOP = Path(os.environ["USERPROFILE"]) / "Desktop"
PROJECT_SRC = Path(__file__).resolve().parent / "src"
PYTHON = sys.executable

SHORTCUTS = [
    {
        "name": "转换工具",
        "target": PYTHON,
        "arguments": str(PROJECT_SRC / "converter_launcher.py"),
        "working_dir": str(PROJECT_SRC),
        "description": "Word↔Markdown 双向智能转换 — 拖放文件到此图标",
    },
    {
        "name": "转换设置",
        "target": PYTHON,
        "arguments": str(PROJECT_SRC / "settings_app.py"),
        "working_dir": str(PROJECT_SRC),
        "description": "配置 Word↔Markdown 转换参数",
    },
]


def create_shortcut(name: str, target: str, arguments: str,
                    working_dir: str, description: str):
    """使用 PowerShell 创建 .lnk，参数通过 base64 编码避免编码问题。"""
    import subprocess
    import base64

    shortcut_path = DESKTOP / f"{name}.lnk"

    # Build PowerShell commands with explicit UTF-8 string handling
    ps_script = f"""
$desktop = [Environment]::GetFolderPath('Desktop')
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut([System.IO.Path]::Combine($desktop, '{name}.lnk'))
$sc.TargetPath = '{target}'
$sc.Arguments = '{arguments}'
$sc.WorkingDirectory = '{working_dir}'
$sc.Description = '{description}'
$sc.Save()
Write-Output 'OK: {name}.lnk'
"""

    # Encode to base64 UTF-16 LE (PowerShell -EncodedCommand expects UTF-16LE)
    ps_bytes = ps_script.encode('utf-16-le')
    b64 = base64.b64encode(ps_bytes).decode('ascii')

    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-EncodedCommand", b64],
        capture_output=True, text=True, timeout=15,
    )
    if result.returncode != 0:
        print(f"FAIL: {name}.lnk — {result.stderr[:200]}")
    else:
        print(result.stdout.strip())


def cleanup_garbled():
    """删除桌面上乱码的 .lnk 文件（修改时间在最近2分钟内的）。"""
    import time
    now = time.time()
    for item in DESKTOP.iterdir():
        if item.suffix == ".lnk":
            age = now - item.stat().st_mtime
            if age < 120:  # 2分钟以内
                # 检查文件名是否包含乱码特征(含无法打印的字符)
                try:
                    item.name.encode('gbk')
                except UnicodeEncodeError:
                    print(f"删除乱码图标: {item.name!r}")
                    item.unlink()
                    continue
                # 也删除我们创建的新图标（以便重建）
                if item.stem in ["转换工具", "转换设置"]:
                    print(f"删除旧图标: {item.name}")
                    item.unlink()


if __name__ == "__main__":
    cleanup_garbled()
    print()
    for cfg in SHORTCUTS:
        create_shortcut(**cfg)
    print("\n完成！请检查桌面。")
