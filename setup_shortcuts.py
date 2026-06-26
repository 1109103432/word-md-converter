"""
桌面快捷方式创建工具 — 开发/测试用。

创建 2 个桌面快捷方式：
  - 开始转换  → converter_launcher.py（拖放文件 + 双击剪贴板→Word）
  - 转换设置  → settings_app.py（参数配置）

用法: python setup_shortcuts.py
     如需删除旧快捷方式再重建: python setup_shortcuts.py --force

工作原理: 通过 base64 编码 PowerShell 命令避免中文字符乱码。
"""
import base64
import os
import subprocess
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parent
SRC = PROJECT / "src"
DESKTOP = Path.home() / "Desktop"

SHORTCUTS = [
    {
        "name": "开始转换",
        "script": str(SRC / "converter_launcher.py"),
        "desc": "Word-Markdown双向转换 - 拖放文件或双击转换剪贴板",
    },
    {
        "name": "转换设置",
        "script": str(SRC / "settings_app.py"),
        "desc": "配置Word-Markdown转换参数",
    },
]


def _run_ps(script: str) -> bool:
    """执行 PowerShell 脚本，base64 编码以避免编码问题。"""
    raw = script.encode("utf-16-le")
    encoded = base64.b64encode(raw).decode()
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-EncodedCommand", encoded],
            capture_output=True, text=True, timeout=15,
        )
        for line in result.stdout.splitlines():
            if line.strip():
                print(f"  {line.strip()}")
        if result.returncode != 0:
            for line in result.stderr.splitlines():
                if line.strip():
                    print(f"  [ERR] {line.strip()}")
            return False
        return True
    except Exception as e:
        print(f"  [FAIL] {e}")
        return False


def main():
    force = "--force" in sys.argv

    pythonw = Path(sys.executable).parent / "pythonw.exe"
    if not pythonw.exists():
        print("  [WARN] pythonw.exe 不存在，回退到 python.exe（会有控制台窗口）")
        pythonw = Path(sys.executable)

    print("创建桌面快捷方式 ...")
    print(f"  项目: {PROJECT}")
    print(f"  Python: {pythonw}")
    print()

    if force:
        # 先删除旧快捷方式
        for sc in SHORTCUTS:
            lnk = DESKTOP / f"{sc['name']}.lnk"
            if lnk.exists():
                lnk.unlink()
                print(f"  [DEL] {sc['name']}.lnk")

    for sc in SHORTCUTS:
        lnk = DESKTOP / f"{sc['name']}.lnk"
        if lnk.exists() and not force:
            print(f"  [SKIP] {sc['name']}.lnk (已存在，--force 可覆盖)")
            continue

        ps = (
            f"$ws = New-Object -ComObject WScript.Shell; "
            f"$sc = $ws.CreateShortcut('{lnk}'); "
            f"$sc.TargetPath = '{pythonw}'; "
            f"$sc.Arguments = '{sc['script']}'; "
            f"$sc.WorkingDirectory = '{DESKTOP}'; "
            f"$sc.Description = '{sc['desc']}'; "
            f"$sc.Save(); "
            f"Write-Output 'Done: {sc['name']}'"
        )
        print(f"  [*] {sc['name']} ...")
        _run_ps(ps)

    print()
    print("完成。图标已创建到桌面。")


if __name__ == "__main__":
    main()
