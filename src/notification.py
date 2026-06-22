"""
Windows 系统通知模块 — 方案 A：PowerShell Toast 子进程。

主进程通过 subprocess.Popen 启动一个独立的 PowerShell 子进程，
子进程负责显示 Windows 原生 Toast 通知并延迟退出。

主进程立即返回，不等待通知完成。子进程独立存活 3 秒，
确保 Windows 通知系统有足够时间拾取并显示 Toast。

要求：开始菜单中需存在 AppUserModelID 快捷方式（由 setup_shortcuts.py 创建）。
"""
import subprocess
import sys
import os
from pathlib import Path

_APP_ID = "WordMD.Converter"
_CREATE_NO_WINDOW = 0x08000000  # 子进程不弹出控制台窗口


def _ensure_start_menu_shortcut():
    """
    确保开始菜单中存在应用快捷方式。

    Windows 10/11 要求：发送带按钮的交互式 Toast 通知的应用
    必须在开始菜单中有对应的快捷方式。无按钮的纯文本通知不需要此步骤。
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
        target = sys.executable
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
            creationflags=_CREATE_NO_WINDOW,
        )
    except Exception:
        pass  # 静默失败，通知仍可工作（仅按钮可能不显示）


# ── 模块加载时执行一次 ──
_ensure_start_menu_shortcut()


def _escape_ps(s: str) -> str:
    """转义字符串中的 PowerShell 特殊字符。"""
    return s.replace("'", "''").replace("\n", "\\n").replace("\r", "")


def show_notification(
    title: str,
    message: str,
    output_path: str = None,
    auto_close="5秒",
    is_error: bool = False,
):
    """
    显示 Windows Toast 通知。

    通过独立 PowerShell 子进程发送，主进程立即返回。
    子进程存活 3 秒确保通知送达。

    受 config.json 中 notification 节控制：
      - enabled=false  → 不显示任何通知
      - show_success=false 且非错误 → 不显示成功通知

    Args:
        title: 通知标题
        message: 通知正文
        output_path: 输出文件路径（显示"打开文件夹"按钮）
        auto_close: "5秒" 或 "25秒"
        is_error: 是否为错误通知
    """
    # ── 通知开关检查 ──
    try:
        from config import load_config
        notif_cfg = load_config().get("notification", {})
        if not notif_cfg.get("enabled", True):
            return
        if not is_error and not notif_cfg.get("show_success", True):
            return
    except Exception:
        pass

    # ── 解析时长 ──
    duration_str = str(auto_close)
    if duration_str in ("5秒", "short", "5"):
        duration = "short"
    else:
        duration = "long"

    # 长通知使用 incomingCall 场景（优先展示）
    scenario = "incomingCall" if duration == "long" else "default"

    # ── 音频 ──
    audio_xml = ""
    if is_error:
        audio_xml = '<audio src="ms-winsoundevent:Notification.Default" />'

    # ── 操作按钮 ──
    actions_xml = ""
    if output_path:
        folder = str(Path(output_path).parent)
        safe_folder = _escape_ps(folder)
        actions_xml = (
            '<actions>'
            f'<action content="📁 打开文件夹" '
            f'arguments="{safe_folder}" '
            f'activationType="protocol" />'
            '</actions>'
        )

    # ── 构建 Toast XML ──
    safe_title = _escape_ps(title)
    safe_msg = _escape_ps(message)

    toast_xml = (
        f'<toast scenario="{scenario}" duration="{duration}">'
        '<visual>'
        '<binding template="ToastGeneric">'
        f'<text>{safe_title}</text>'
        f'<text>{safe_msg}</text>'
        '</binding>'
        '</visual>'
        f'{actions_xml}'
        f'{audio_xml}'
        '</toast>'
    )

    # ── PowerShell 脚本 ──
    # 在 PowerShell 行内用双引号包裹 XML，XML 内部的双引号已由
    # _escape_ps 处理（通过 '' 转义单引号，这里 XML 属性用单引号）。
    # 改用 here-string 避免嵌套引号问题。
    ps_script = f'''
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml(@'
{toast_xml}
'@)
$toast = New-Object Windows.UI.Notifications.ToastNotification($xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("{_APP_ID}").Show($toast)
Start-Sleep -Seconds 3
'''

    # ── 启动独立子进程 ──
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-Command", ps_script],
            creationflags=_CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass  # 通知失败不阻断主流程
