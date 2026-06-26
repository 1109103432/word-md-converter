"""
Word-MD快速转换 — 统一转换入口
用法：将 .docx/.doc 或 .md 文件拖放到桌面"开始转换"图标上即可自动识别并转换。
     支持同时拖入多个文件。
     或命令行：python converter_launcher.py <文件1> <文件2> ...

自动判断规则：
  - .docx → 转换为 .md (Markdown)
  - .doc  → 通过 Word COM 转为 .docx 后再转 .md（需本机安装 Word）
  - .md / .markdown / .txt / .text → 转换为 .docx (Word)

支持 PyInstaller 打包和直接源码运行两种模式。
"""
import subprocess
import sys
import shutil
from pathlib import Path

# 源码运行时确保能找到同目录下的模块
# PyInstaller 打包后模块已内置，无需此操作
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from converter import docx_to_markdown, markdown_to_docx
from notification import show_notification
from config import load_config

# ── 支持的文件类型 ──
WORD_EXTENSIONS = {".docx"}                        # python-docx 原生支持
LEGACY_WORD_EXTENSIONS = {".doc"}                  # 旧格式，Word COM 中转
MD_EXTENSIONS = {".md", ".markdown", ".txt", ".text"}

# ── Pandoc 路径缓存 ──
_PANDOC_PATH: str | None = None


def _get_pandoc_exe() -> str | None:
    """查找 pandoc 可执行文件，结果缓存。"""
    global _PANDOC_PATH
    if _PANDOC_PATH is not None:
        return _PANDOC_PATH

    # 先尝试 pandoc_engine 的查找逻辑
    try:
        from pandoc_engine import _find_pandoc
        _PANDOC_PATH = _find_pandoc()
        if _PANDOC_PATH:
            return _PANDOC_PATH
    except ImportError:
        pass

    # 回退到系统 PATH
    found = shutil.which("pandoc")
    _PANDOC_PATH = found if found else ""
    return _PANDOC_PATH or None


def main():
    config = load_config()
    notif = config.get("notification", {})
    duration = notif.get("duration", "short")

    if len(sys.argv) < 2:
        _convert_clipboard_to_docx(config, duration)
        return

    # ── 收集并分类所有输入文件 ──
    word_files = []
    doc_files = []    # 旧格式 .doc，单独处理
    md_files = []
    not_found = []
    unsupported = []

    for arg in sys.argv[1:]:
        p = Path(arg)
        if not p.exists():
            not_found.append(p)
            continue
        suffix = p.suffix.lower()
        if suffix in WORD_EXTENSIONS:
            word_files.append(p)
        elif suffix in LEGACY_WORD_EXTENSIONS:
            doc_files.append(p)
        elif suffix in MD_EXTENSIONS:
            md_files.append(p)
        else:
            unsupported.append(p)

    # ── 统计 ──
    ok, fail = 0, 0
    total = len(word_files) + len(doc_files) + len(md_files) + len(not_found) + len(unsupported)
    batch_mode = total > 1

    for p in word_files:
        if _convert_word_to_md(p, config, duration, notify=not batch_mode):
            ok += 1
        else:
            fail += 1

    for p in doc_files:
        if _convert_doc_to_md(p, config, duration, notify=not batch_mode):
            ok += 1
        else:
            fail += 1

    for p in md_files:
        if _convert_md_to_word(p, config, duration, notify=not batch_mode):
            ok += 1
        else:
            fail += 1

    # ── 汇总通知 ──
    if total == 0:
        return
    if total == 1 and ok + fail > 0:
        # 单文件且已处理：通知已在上方显示，无需汇总
        return
    if total == 1 and not_found:
        show_notification(
            title="转换失败",
            message=f"找不到文件：\n{not_found[0]}",
            is_error=True,
            auto_close=10,
        )
        return
    if total == 1 and unsupported:
        p = unsupported[0]
        show_notification(
            title="格式不支持",
            message=f"无法识别该文件类型。\n\n"
                    f"文件：{p.name}\n"
                    f"后缀：{p.suffix}\n\n"
                    f"支持：.docx / .doc / .md / .markdown / .txt",
            is_error=True,
            auto_close=10,
        )
        return

    parts = [f"✅ 成功：{ok} 个"]
    if fail:
        parts.append(f"❌ 失败：{fail} 个")
    if not_found:
        parts.append(f"⚠ 文件不存在：{len(not_found)} 个")
    if unsupported:
        names = ", ".join(p.name for p in unsupported[:3])
        extra = "..." if len(unsupported) > 3 else ""
        parts.append(f"⊘ 格式不支持：{names}{extra}")

    is_error = fail > 0 or not_found or unsupported
    show_notification(
        title="📦 批量转换完成" if not is_error else "📦 批量转换完成（有问题）",
        message="\n".join(parts),
        is_error=is_error,
        auto_close=duration if not is_error else "long",
    )


def _convert_word_to_md(input_path: Path, config: dict, duration: str,
                        notify: bool = True) -> bool:
    """Word → Markdown 转换流程。返回 True=成功, False=失败。"""
    try:
        output_path = input_path.with_suffix(".md")

        md_content = docx_to_markdown(str(input_path), config, output_path)

        with open(output_path, "w", encoding=config.get("output_encoding", "utf-8")) as f:
            f.write(md_content)

        if notify:
            show_notification(
                title="✅ 转换成功  Word → Markdown",
                message=f"源文件：{input_path.name}\n"
                        f"输出文件：{output_path.name}",
                output_path=str(output_path),
                auto_close=duration,
            )
        return True

    except Exception as e:
        if notify:
            show_notification(
                title="转换失败  Word → Markdown",
                message=f"转换过程中发生错误：\n{str(e)}",
                is_error=True,
                auto_close="long",
            )
        return False


def _convert_doc_to_md(input_path: Path, config: dict, duration: str,
                       notify: bool = True) -> bool:
    """旧格式 .doc → Markdown 转换（Word COM 后台转为 .docx 再转换）。

    .doc 是 Word 97-2003 二进制格式，Pandoc 和 python-docx 均不支持。
    利用本机已安装的 Microsoft Word 将其转为 .docx，再走正常转换流程。
    返回 True=成功, False=失败。
    """
    import tempfile

    # ── 1. Word COM 将 .doc 转为 .docx ──
    temp_dir = Path(tempfile.gettempdir())
    temp_docx = temp_dir / f"_wdmd_{input_path.stem}.docx"

    try:
        ps_script = (
            f'$word = New-Object -ComObject Word.Application; '
            f'$word.Visible = $false; '
            f'$word.DisplayAlerts = 0; '
            f'$doc = $word.Documents.Open(\'{input_path}\'); '
            f'$doc.SaveAs2(\'{temp_docx}\', 16); '
            f'$doc.Close(); '
            f'$word.Quit(); '
            f'[Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null'
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-Command", ps_script],
            capture_output=True, text=True, timeout=60,
            creationflags=0x08000000 if sys.platform == "win32" else 0,
        )

        if result.returncode != 0 or not temp_docx.exists():
            if notify:
                show_notification(
                    title="❌ .doc 转换失败 · 需安装 Word",
                    message=f"文件：{input_path.name}\n\n"
                            f".doc 是旧版 Word 格式（97-2003），\n"
                            f"Pandoc 和 python-docx 均不支持。\n\n"
                            f"💡 解决方式：\n"
                            f"  1. 用 Word 打开该文件\n"
                            f"  2. 另存为 .docx 格式\n"
                            f"  3. 重新拖入 .docx 文件",
                    is_error=True,
                    auto_close="long",
                )
            return False

    except (subprocess.TimeoutExpired, Exception):
        if notify:
            show_notification(
                title="❌ .doc 转换失败",
                message=f"文件：{input_path.name}\n"
                        f"Word 后台转换时出错。\n\n"
                        f"💡 建议用 Word 打开文件，另存为\n"
                        f"  .docx 格式后重新拖入。",
                is_error=True,
                auto_close="long",
            )
        return False

    # ── 2. 对 .docx 执行正常 Word→MD 转换 ──
    try:
        ok = _convert_word_to_md(temp_docx, config, duration, notify=False)
        if not ok:
            if notify:
                show_notification(
                    title="❌ .doc 转换失败",
                    message=f"文件：{input_path.name}\n\n"
                            f"已通过 Word 转为 .docx，但后续转换失败。\n"
                            f"💡 建议直接用 Word 另存为 .docx 后重新拖入。",
                    is_error=True,
                    auto_close="long",
                )
            return False

        # 输出路径修正为原文件路径的 .md
        output_path = input_path.with_suffix(".md")
        temp_md = temp_docx.with_suffix(".md")
        if temp_md.exists() and temp_md != output_path:
            output_path.write_text(temp_md.read_text(encoding="utf-8"),
                                   encoding="utf-8")
            temp_md.unlink()

        if notify:
            show_notification(
                title="⚠ .doc 已转换 · 格式已过时",
                message=f"源文件：{input_path.name}\n"
                        f"输出文件：{output_path.name}\n\n"
                        f"⚠ .doc 是 Word 97-2003 旧格式。\n"
                        f"  建议用 Word 另存为 .docx 后重试，\n"
                        f"  可获得更稳定效果。",
                output_path=str(output_path),
                auto_close="long",
            )
        return True

    except Exception as e:
        if notify:
            show_notification(
                title="❌ .doc 转换失败",
                message=f"文件：{input_path.name}\n"
                        f"错误：{str(e)}\n\n"
                        f"💡 建议用 Word 另存为 .docx 后重试。",
                is_error=True,
                auto_close="long",
            )
        return False

    finally:
        if temp_docx.exists():
            try:
                temp_docx.unlink()
            except OSError:
                pass


def _convert_md_to_word(input_path: Path, config: dict, duration: str,
                        notify: bool = True) -> bool:
    """Markdown → Word 转换流程。返回 True=成功, False=失败。"""
    try:
        output_path = input_path.with_suffix(".docx")

        with open(input_path, "r", encoding="utf-8") as f:
            md_content = f.read()

        markdown_to_docx(md_content, str(output_path), config)

        if notify:
            show_notification(
                title="✅ 转换成功  Markdown → Word",
                message=f"源文件：{input_path.name}\n"
                        f"输出文件：{output_path.name}",
                output_path=str(output_path),
                auto_close=duration,
            )
        return True

    except Exception as e:
        if notify:
            show_notification(
                title="转换失败  Markdown → Word",
                message=f"转换过程中发生错误：\n{str(e)}",
                is_error=True,
                auto_close="long",
            )
        return False


def _convert_clipboard_to_docx(config: dict, duration: str):
    """双击图标时：读取剪贴板内容 → 识别为 Markdown → 输出 Word 文档。

    输出位置：exe 所在目录（打包模式）或项目根目录（开发模式）。
    文件名根据剪贴板首行标题自动生成，无标题时用时间戳。
    """
    import tkinter as tk
    from datetime import datetime

    # ── 1. 读取剪贴板 ──
    try:
        root = tk.Tk()
        root.withdraw()
        md_content = root.clipboard_get()
        root.destroy()
    except tk.TclError:
        # 剪贴板无文本内容
        show_notification(
            title="Word-MD快速转换",
            message="📋 剪贴板内无文本内容。\n\n"
                    "💡 双击图标 = 剪贴板 → Word\n"
                    "💡 拖入文件 = 自动识别转换\n\n"
                    "📄  拖入 .docx → .md\n"
                    "📄  拖入 .doc → .md（旧格式，建议先转 .docx）\n"
                    "📝  拖入 .md → .docx",
            auto_close=8,
        )
        return
    except Exception as e:
        show_notification(
            title="剪贴板读取失败",
            message=f"无法读取剪贴板：\n{str(e)}",
            is_error=True,
            auto_close="long",
        )
        return

    if not md_content or not md_content.strip():
        show_notification(
            title="Word-MD快速转换",
            message="📋 剪贴板为空。\n\n"
                    "💡 双击图标 = 剪贴板 → Word\n"
                    "💡 拖入文件 = 自动识别转换",
            auto_close=8,
        )
        return

    # ── 2. 确定输出目录（快捷方式所在目录 = 当前工作目录）──
    output_dir = Path.cwd()

    # ── 3. 生成文件名（优先用 md 首行标题）──
    first_line = md_content.strip().split("\n")[0].strip()
    title = first_line.lstrip("#").strip()
    if title and len(title) <= 80:
        # 去除文件名非法字符
        safe_title = "".join(c for c in title if c not in r'\/:*?"<>|')
        safe_title = safe_title.strip()
        if safe_title:
            filename = f"{safe_title}.docx"
        else:
            filename = f"剪贴板_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    else:
        filename = f"剪贴板_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

    output_path = output_dir / filename

    # ── 4. 避免覆盖已有文件 ──
    if output_path.exists():
        stem = output_path.stem
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = output_dir / f"{stem}_{ts}.docx"

    # ── 5. 转换 ──
    try:
        markdown_to_docx(md_content, str(output_path), config)
        show_notification(
            title="✅ 剪贴板 → Word 转换成功",
            message=f"📋 剪贴板内容已转为 Word 文档\n\n"
                    f"📄 {output_path.name}\n\n"
                    f"📁 {output_dir}",
            output_path=str(output_path),
            auto_close=duration,
        )
    except Exception as e:
        show_notification(
            title="剪贴板转换失败",
            message=f"转换过程中发生错误：\n{str(e)}",
            is_error=True,
            auto_close="long",
        )


if __name__ == "__main__":
    main()
