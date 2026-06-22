"""
Windows 系统通知模块 — 使用 win11toast (WinRT 原生 API)。

win11toast 在进程内直接调用 Windows Runtime ToastNotification API，
不依赖子进程。默认使用 'Python' 作为 AppUserModelID（Python 安装时
已注册，无需额外配置），兼容 pythonw.exe 后台运行。

流程：
  notify() → WinRT Show() → sleep(0.5s) → 主进程退出
                                    ↑
                          Windows 通知系统已拾取 Toast
"""
import sys
import os
import time
from pathlib import Path

from win11toast import notify as _win11_notify

# 使用 Python 默认 AppUserModelID，无需开始菜单快捷方式
_APP_ID = "Python"


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
        output_path: （暂未使用 — 系统限制：自定义 AppID 需要额外注册）
        auto_close: "5秒" 或 "25秒"
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

    # ── 时长 ──
    duration = "short" if str(auto_close) in ("5秒", "short", "5") else "long"

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
            duration=duration,
            audio=audio,
        )
        # WinRT Show() 是异步的 — 短暂等待确保 Windows 拾取通知
        time.sleep(0.5)
    except Exception:
        pass  # 通知静默失败，不阻断主流程
