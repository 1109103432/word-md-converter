"""
Word ↔ Markdown 统一转换入口
用法：将 .docx 或 .md 文件拖放到桌面"转换工具"图标上即可自动识别并转换。
     或命令行：python converter_launcher.py <文件路径>

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
        # 没有拖入文件 — 显示使用提示
        show_notification(
            title="Word ↔ Markdown 转换工具",
            message="请将文件拖放到此图标上即可自动转换。\n\n"
                    "📄  拖入 .docx → 自动转为 .md\n"
                    "📝  拖入 .md   → 自动转为 .docx\n\n"
                    "输出位置：原文件所在目录",
            auto_close=8,
        )
        return

    input_path = Path(sys.argv[1])

    # 文件存在性检查
    if not input_path.exists():
        show_notification(
            title="转换失败",
            message=f"找不到文件：\n{input_path}",
            is_error=True,
            auto_close=10,
        )
        return

    suffix = input_path.suffix.lower()

    # ── 判断转换方向 ──
    if suffix in WORD_EXTENSIONS:
        _convert_word_to_md(input_path, config, duration)
    elif suffix in MD_EXTENSIONS:
        _convert_md_to_word(input_path, config, duration)
    else:
        show_notification(
            title="格式不支持",
            message=f"无法识别该文件类型。\n\n"
                    f"当前文件后缀：{suffix}\n"
                    f"支持的格式：\n"
                    f"  📄 Word → MD：.docx\n"
                    f"  📝 MD → Word：.md, .markdown, .txt",
            is_error=True,
            auto_close=10,
        )


def _convert_word_to_md(input_path: Path, config: dict, duration: str):
    """Word → Markdown 转换流程。"""
    try:
        output_path = input_path.with_suffix(".md")

        md_content = docx_to_markdown(str(input_path), config, output_path)

        # Pandoc 路径已直接写入文件；此处覆盖以确保编码一致
        with open(output_path, "w", encoding=config.get("output_encoding", "utf-8")) as f:
            f.write(md_content)

        show_notification(
            title="✅ 转换成功  Word → Markdown",
            message=f"源文件：{input_path.name}\n"
                    f"输出文件：{output_path.name}",
            output_path=str(output_path),
            auto_close=duration,
        )

    except Exception as e:
        show_notification(
            title="转换失败  Word → Markdown",
            message=f"转换过程中发生错误：\n{str(e)}",
            is_error=True,
            auto_close="long",
        )


def _convert_md_to_word(input_path: Path, config: dict, duration: str):
    """Markdown → Word 转换流程。"""
    try:
        output_path = input_path.with_suffix(".docx")

        with open(input_path, "r", encoding="utf-8") as f:
            md_content = f.read()

        markdown_to_docx(md_content, str(output_path), config)

        show_notification(
            title="✅ 转换成功  Markdown → Word",
            message=f"源文件：{input_path.name}\n"
                    f"输出文件：{output_path.name}",
            output_path=str(output_path),
            auto_close=duration,
        )

    except Exception as e:
        show_notification(
            title="转换失败  Markdown → Word",
            message=f"转换过程中发生错误：\n{str(e)}",
            is_error=True,
            auto_close="long",
        )


if __name__ == "__main__":
    main()
