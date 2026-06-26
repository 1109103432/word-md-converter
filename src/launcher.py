"""
Word-MD快速转换 — 统一启动入口。

用法:
    转换工具.exe                 → 双击：剪贴板→Word；拖入文件：自动识别转换
    转换工具.exe --settings      → 打开参数设置窗口
    转换工具.exe <文件1> <文件2>  → 转换拖入的文件

两个桌面快捷方式指向同一个 exe：
    开始转换.lnk  → 转换工具.exe
    转换设置.lnk  → 转换工具.exe --settings

此设计将两个 exe 合并为一个，去除重复的 Qt 框架（节省 ~56 MB）。
"""
import sys
from pathlib import Path

# 源码运行时确保能找到同目录下的模块
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parent))


def main():
    # ── 检测启动模式 ──
    if "--settings" in sys.argv:
        # 移除 --settings 避免干扰 argparse
        sys.argv.remove("--settings")
        _run_settings()
    else:
        _run_converter()


def _run_converter():
    """启动转换器模式（拖放文件 / 剪贴板→Word）。"""
    from converter_launcher import main as converter_main
    converter_main()


def _run_settings():
    """启动参数设置窗口。"""
    from settings_app import main as settings_main
    settings_main()


if __name__ == "__main__":
    main()
