"""
Windows 系统通知模块 — 使用 win11toast (WinRT 原生 API)。

通过自定义 AppUserModelID + 开始菜单快捷方式，支持：
- 点击通知打开输出文件所在文件夹
- 多种通知停留时长（短/中/长/持续）

流程：
  notify() → WinRT Show() → sleep(0.5s) → 退出
"""
import os
import subprocess
import sys
import time
from pathlib import Path

from win11toast import notify as _win11_notify

_APP_ID = "WordMD.Converter"


def _ensure_start_menu_shortcut():
    """
    确保开始菜单中存在应用快捷方式。

    Windows 要求：带交互操作（如点击打开文件）的 Toast 通知
    必须在开始菜单中有对应 AppUserModelID 的快捷方式。
    """
    shortcut_dir = (
        Path(os.environ["APPDATA"])
        / "Microsoft" / "Windows" / "Start Menu" / "Programs"
    )
    shortcut_dir.mkdir(parents=True, exist_ok=True)
    shortcut_path = shortcut_dir / f"{_APP_ID}.lnk"

    if shortcut_path.exists():
        return

    if getattr(sys, 'frozen', False):
        target = sys.executable
        working_dir = str(Path(sys.executable).parent)
    else:
        # 开发模式：用 pythonw.exe 避免控制台
        target = str(Path(sys.executable).parent / "pythonw.exe")
        working_dir = str(Path(__file__).resolve().parent.parent)

    ps = (
        f"$ws = New-Object -ComObject WScript.Shell; "
        f"$sc = $ws.CreateShortcut('{shortcut_path}'); "
        f"$sc.TargetPath = '{target}'; "
        f"$sc.WorkingDirectory = '{working_dir}'; "
        f"$sc.Save(); "
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-Command", ps],
            capture_output=True, timeout=10,
            creationflags=0x08000000,
        )
    except Exception:
        pass  # 静默失败，通知仍可显示（仅点击操作不可用）


# ── 模块加载时注册 ──
_ensure_start_menu_shortcut()


# ── 时长映射：用户配置 → Toast 参数 ──
_DURATION_MAP = {
    # 短 (系统 ~7s)
    "5秒":   {"duration": "short", "scenario": "default"},
    "short": {"duration": "short", "scenario": "default"},
    "5":     {"duration": "short", "scenario": "default"},
    # 中 (~12s, 用 long+default)
    "10秒":  {"duration": "long",  "scenario": "default"},
    "medium": {"duration": "long",  "scenario": "default"},
    "10":    {"duration": "long",  "scenario": "default"},
    # 长 (~25s)
    "25秒":  {"duration": "long",  "scenario": "default"},
    "long":  {"duration": "long",  "scenario": "default"},
    "25":    {"duration": "long",  "scenario": "default"},
    # 持续显示（手动关闭才消失）
    "持续":   {"duration": "long",  "scenario": "reminder"},
    "keep":   {"duration": "long",  "scenario": "reminder"},
}


def show_notification(
    title: str,
    message: str,
    output_path: str = None,
    auto_close="5秒",
    is_error: bool = False,
):
    """
    显示 Windows 10/11 系统原生 Toast 通知。

    Args:
        title: 通知标题
        message: 通知正文
        output_path: 输出文件路径（点击通知打开所在文件夹）
        auto_close: 停留时长 —
            "5秒"/"short"  → 短 (~7s)
            "10秒"/"medium" → 中 (~12s)
            "25秒"/"long"   → 长 (~25s)
            "持续"/"keep"   → 持续显示（手动关闭）
        is_error: 是否为错误通知
    """
    # ── 通知开关 ──
    try:
        from config import load_config
        notif_cfg = load_config().get("notification", {})
        if not notif_cfg.get("enabled", True):
            return
        if not is_error and not notif_cfg.get("show_success", True):
            return
    except Exception:
        pass

    # ── 时长解析 ──
    key = str(auto_close)
    params = _DURATION_MAP.get(key, _DURATION_MAP["5秒"])

    # ── 点击操作：打开输出文件 ──
    on_click = None
    if output_path:
        # 传入文件路径，Windows 用默认程序打开
        on_click = str(Path(output_path).resolve())

    # ── 音频 ──
    audio = None
    if is_error:
        audio = "ms-winsoundevent:Notification.Default"

    # ── 发送通知 ──
    try:
        _win11_notify(
            app_id=_APP_ID,
            title=title,
            body=message,
            duration=params["duration"],
            scenario=params["scenario"],
            on_click=on_click,
            audio=audio,
        )
        time.sleep(0.5)
    except Exception:
        pass
