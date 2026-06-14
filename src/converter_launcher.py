"""
Word ↔ Markdown 统一转换入口
用法：将 .docx 或 .md 文件拖放到桌面"转换工具"图标上即可自动识别并转换。
     支持同时拖入多个文件。
     或命令行：python converter_launcher.py <文件1> <文件2> ...

自动判断规则：
  - .docx → 转换为 .md (Markdown)
  - .md / .markdown / .txt / .text → 转换为 .docx (Word)

支持 PyInstaller 打包和直接源码运行两种模式。
"""
import sys
from pathlib import Path

# 源码运行时确保能找到同目录下的模块
# PyInstaller 打包后模块已内置，无需此操作
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from converter import docx_to_markdown, markdown_to_docx
from notification import show_notification
from config import load_config

# ── 支持的文件类型 ──
WORD_EXTENSIONS = {".docx"}  # python-docx 仅支持 .docx
MD_EXTENSIONS = {".md", ".markdown", ".txt", ".text"}


def main():
    config = load_config()
    notif = config.get("notification", {})
    duration = notif.get("duration", "short")

    if len(sys.argv) < 2:
        show_notification(
            title="Word ↔ Markdown 转换工具",
            message="请将文件拖放到此图标上即可自动转换。\n\n"
                    "📄  拖入 .docx → 自动转为 .md\n"
                    "📝  拖入 .md   → 自动转为 .docx\n\n"
                    "📦  支持同时拖入多个文件\n"
                    "输出位置：原文件所在目录",
            auto_close=8,
        )
        return

    # ── 收集并分类所有输入文件 ──
    word_files = []   # (Path, suffix)
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
        elif suffix in MD_EXTENSIONS:
            md_files.append(p)
        else:
            unsupported.append(p)

    # ── 统计 ──
    ok, fail = 0, 0
    total = len(word_files) + len(md_files) + len(not_found) + len(unsupported)
    batch_mode = total > 1

    for p in word_files:
        if _convert_word_to_md(p, config, duration, notify=not batch_mode):
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
                    f"支持：.docx / .md / .markdown / .txt",
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


if __name__ == "__main__":
    main()
