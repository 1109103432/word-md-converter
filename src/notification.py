"""
Windows 系统通知模块 — 通过 Windows Toast 通知显示转换结果。

优先使用 winotify 发送系统原生通知（进入操作中心），
如不可用则回退到 tkinter 浮窗方案。

Toast 通知特性:
- 自动进入 Windows 操作中心，可回溯查看
- 支持操作按钮（如"打开文件夹"）
- 非阻塞：通知弹出后脚本立即返回
- 持续时间: 5秒 / 25秒
"""
import ctypes
import os
import subprocess
import sys
from pathlib import Path

_APP_ID = "WordMD.Converter"

# ── 注册 AppUserModelID（Windows Toast 通知的必要前提）──
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(_APP_ID)
except Exception:
    pass


def _ensure_start_menu_shortcut():
    """
    确保开始菜单中存在应用快捷方式。

    Windows 10/11 要求：发送交互式 Toast 通知（含操作按钮）的应用
    必须在开始菜单中有对应的快捷方式，否则通知会被静默丢弃。

    在 PyInstaller 打包模式下指向 .exe 自身，开发模式下指向 python.exe。
    """
    shortcut_dir = (
        Path(os.environ["APPDATA"])
        / "Microsoft" / "Windows" / "Start Menu" / "Programs"
    )
    shortcut_dir.mkdir(parents=True, exist_ok=True)
    shortcut_path = shortcut_dir / f"{_APP_ID}.lnk"

    if shortcut_path.exists():
        return  # 已存在，跳过

    # 确定目标可执行文件
    if getattr(sys, 'frozen', False):
        target = sys.executable
        working_dir = str(Path(sys.executable).parent)
    else:
        target = sys.executable
        working_dir = str(Path(__file__).resolve().parent.parent)

    # 使用 PowerShell 创建快捷方式
    ps_cmd = (
        f"$ws = New-Object -ComObject WScript.Shell; "
        f"$sc = $ws.CreateShortcut('{shortcut_path}'); "
        f"$sc.TargetPath = '{target}'; "
        f"$sc.WorkingDirectory = '{working_dir}'; "
        f"$sc.Save(); "
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-Command", ps_cmd],
            capture_output=True, timeout=10,
        )
    except Exception:
        pass  # 静默失败，通知可能仍可工作（仅无按钮时）


# ── 模块加载时执行一次 ──
_ensure_start_menu_shortcut()

# ── 尝试导入 winotify ──
try:
    from winotify import Notification, audio
    _HAS_WINOTIFY = True
except ImportError:
    _HAS_WINOTIFY = False


def show_notification(
    title: str,
    message: str,
    output_path: str = None,
    auto_close = "5秒",
    is_error: bool = False,
):
    """
    显示 Windows Toast 通知。

    受 config.json 中 notification 节控制：
      - enabled=false  → 不显示任何通知
      - show_success=false 且非错误 → 不显示成功通知

    Args:
        title: 通知标题
        message: 通知正文
        output_path: 输出文件路径（会显示"打开文件夹"按钮）
        auto_close: 通知持续时间，"5秒" (~5s) 或 "25秒" (~25s)
        is_error: 是否为错误通知
    """
    # ── 通知开关检查 ──
    try:
        from config import load_config
        notif_cfg = load_config().get("notification", {})
        if not notif_cfg.get("enabled", True):
            return  # 全局关闭
        if not is_error and not notif_cfg.get("show_success", True):
            return  # 仅关闭成功通知
    except Exception:
        pass  # 配置读取失败时照常显示通知

    if _HAS_WINOTIFY:
        _show_toast(title, message, output_path, auto_close, is_error)
    else:
        _show_tk_fallback(title, message, output_path, auto_close, is_error)


def _parse_duration(value: str) -> str:
    """将用户友好的时长字符串转换为 winotify 格式。

    "5秒" / "short" / "5" → "short"
    "25秒" / "long" / "25" → "long"
    """
    if value in ("5秒", "short", "5"):
        return "short"
    return "long"


def _parse_duration_seconds(value) -> int:
    """将任何时长格式转换为秒数（供 tkinter 回退方案使用）。"""
    if isinstance(value, str):
        if value in ("5秒", "short", "5"):
            return 5
        return 25
    return int(value) if value else 5


def _show_toast(
    title: str,
    message: str,
    output_path: str = None,
    auto_close = "5秒",
    is_error: bool = False,
):
    """使用 winotify 发送系统 Toast 通知。"""
    duration = _parse_duration(auto_close) if isinstance(auto_close, str) else (
        "long" if (auto_close > 5 or auto_close == 0) else "short"
    )

    # 构建通知
    toast = Notification(
        app_id=_APP_ID,
        title=title,
        msg=message,
        duration=duration,
        icon=None,  # 使用系统默认图标
    )

    # 错误通知使用警告音效
    if is_error:
        toast.set_audio(audio.Default, loop=False)

    # 输出路径 → 操作按钮
    if output_path:
        folder = str(Path(output_path).parent)
        # winotify 的 add_actions 接受 (label, launch) 对
        # launch 可以是任意字符串；我们在回调中处理
        toast.add_actions(label="📁 打开文件夹", launch=folder)

    toast.show()


def _show_tk_fallback(
    title: str,
    message: str,
    output_path: str = None,
    auto_close = "5秒",
    is_error: bool = False,
):
    """回退方案：使用 tkinter 浮窗通知。"""
    seconds = _parse_duration_seconds(auto_close)
    import tkinter as tk

    root = tk.Tk()
    root.title("")
    root.overrideredirect(True)
    root.attributes("-topmost", True)

    width, height = 360, 200 if output_path else 160

    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = screen_w - width - 20
    y = screen_h - height - 40
    root.geometry(f"{width}x{height}+{x}+{y}")

    if is_error:
        accent_color = "#E74C3C"
        bg_color = "#FDEDEC"
    else:
        accent_color = "#27AE60"
        bg_color = "#EAFAF1"

    root.configure(bg=bg_color)

    main = tk.Frame(root, bg=bg_color, padx=16, pady=12)
    main.pack(fill=tk.BOTH, expand=True)

    bar = tk.Frame(main, bg=accent_color, height=4)
    bar.pack(fill=tk.X, pady=(0, 10))

    header = tk.Frame(main, bg=bg_color)
    header.pack(fill=tk.X)

    icon_text = "✕" if is_error else "✓"
    tk.Label(
        header, text=icon_text,
        fg="white", bg=accent_color,
        font=("Arial", 14, "bold"), width=2, height=1,
    ).pack(side=tk.LEFT, padx=(0, 8))

    tk.Label(
        header, text=title,
        fg="#2C3E50", bg=bg_color,
        font=("Microsoft YaHei", 12, "bold"),
    ).pack(side=tk.LEFT)

    close_btn = tk.Label(
        header, text="✕", fg="#95A5A6", bg=bg_color,
        font=("Arial", 12), cursor="hand2",
    )
    close_btn.pack(side=tk.RIGHT)
    close_btn.bind("<Button-1>", lambda e: root.destroy())

    tk.Label(
        main, text=message,
        fg="#555", bg=bg_color,
        font=("Microsoft YaHei", 10),
        justify=tk.LEFT, anchor="w", wraplength=width - 40,
    ).pack(fill=tk.X, pady=(4, 6))

    if output_path:
        path_frame = tk.Frame(main, bg="#F0F3F4", bd=1, relief=tk.SOLID)
        path_frame.pack(fill=tk.X, pady=(4, 2))

        display_path = output_path
        if len(display_path) > 50:
            display_path = "..." + display_path[-47:]

        path_label = tk.Label(
            path_frame, text=f"📁  {display_path}",
            fg="#2980B9", bg="#F0F3F4",
            font=("Consolas", 9), cursor="hand2", anchor="w",
        )
        path_label.pack(fill=tk.X, padx=8, pady=4)
        path_label.bind("<Button-1>", lambda e: os.startfile(str(Path(output_path).parent)))

    if seconds > 0:
        root.after(seconds * 1000, root.destroy)

    root.bind("<Button-1>", lambda e: root.destroy())
    root.mainloop()
