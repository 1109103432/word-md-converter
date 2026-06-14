"""
Pandoc 转换引擎 — 调用 Pandoc 将 Markdown 转换为 Word 文档。

Pandoc 查找顺序:
  1. 当前可执行文件/脚本所在目录（便携版捆绑）
  2. 项目根目录
  3. 系统 PATH

如未找到 Pandoc，自动回退到 python-docx 原生实现。
"""
import subprocess
import sys
import shutil
from pathlib import Path


# ── Pandoc 路径缓存 ──
_PANDOC_PATH: str | None = None
_PANDOC_SEARCHED = False


def _find_pandoc() -> str | None:
    """查找 pandoc 可执行文件，返回路径或 None。结果会被缓存。"""
    global _PANDOC_PATH, _PANDOC_SEARCHED
    if _PANDOC_SEARCHED:
        return _PANDOC_PATH
    _PANDOC_SEARCHED = True

    # 1. 确定应用所在目录
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后
        app_dir = Path(sys.executable).resolve().parent
    else:
        # 开发模式：本文件在 src/ 下，项目根在 src/ 父级
        app_dir = Path(__file__).resolve().parent.parent

    # 2. 按优先级搜索
    search_dirs = [
        app_dir,                        # exe 同目录（便携版首选）
        app_dir.parent,                 # 上级目录（dist/ 场景）
    ]

    for base in search_dirs:
        for exe_name in ("pandoc.exe", "pandoc"):
            candidate = base / exe_name
            if candidate.is_file():
                _PANDOC_PATH = str(candidate)
                return _PANDOC_PATH

    # 3. 系统 PATH
    found = shutil.which("pandoc")
    if found:
        _PANDOC_PATH = found
        return _PANDOC_PATH

    return None


def has_pandoc() -> bool:
    """检测是否可调用 Pandoc。"""
    return _find_pandoc() is not None


def get_pandoc_version() -> str | None:
    """获取 Pandoc 版本号，如 '3.9.0.2'。"""
    exe = _find_pandoc()
    if not exe:
        return None
    try:
        result = subprocess.run(
            [exe, "--version"],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode == 0:
            return result.stdout.strip().split("\n")[0].split()[-1]
    except Exception:
        pass
    return None


def _get_app_dir() -> Path:
    """获取应用根目录（与 config._get_app_dir 一致的逻辑）。"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    else:
        return Path(__file__).resolve().parent.parent


def get_builtin_template_path() -> Path | None:
    """返回内置模板 template.docx 的路径，不存在则返回 None。"""
    template = _get_app_dir() / "template.docx"
    if template.is_file():
        return template
    return None


def resolve_reference_doc(value: str = None) -> str | None:
    """
    将配置中的 pandoc_reference_doc 值解析为实际文件路径。

    解析规则：
      - "built-in" 或空字符串 → 查找内置 template.docx
      - 其他值 → 作为文件路径，存在则使用

    返回 Path 或 None（表示不使用模板）。
    """
    if not value or value == "built-in":
        # 使用内置模板
        template = get_builtin_template_path()
        return str(template) if template else None

    # 自定义路径
    path = Path(value)
    if path.is_file():
        return str(path)

    # 路径无效，回退到内置模板
    template = get_builtin_template_path()
    return str(template) if template else None


def markdown_to_docx_pandoc(
    md_text: str,
    output_path: str | Path,
    reference_doc: str | Path = None,
) -> bool:
    """
    通过 Pandoc 将 Markdown 转换为 Word 文档。

    Args:
        md_text: Markdown 原文
        output_path: 输出 .docx 路径
        reference_doc: 可选，参考样式模板 .docx。
                      支持 "built-in"（使用内置 template.docx）、
                      自定义路径、或 None（不使用模板）

    Returns:
        True 表示转换成功，False 表示失败
    """
    exe = _find_pandoc()
    if not exe:
        return False

    # 解析模板路径
    resolved_ref = resolve_reference_doc(
        str(reference_doc) if reference_doc else None
    )

    args = [
        exe,
        "--from=markdown+hard_line_breaks+pipe_tables+fenced_code_blocks",
        "--to=docx",
        f"--output={output_path}",
        # 嵌入图片数据而非引用外部文件
        "--embed-resources",
    ]

    if resolved_ref:
        args.append(f"--reference-doc={resolved_ref}")

    try:
        result = subprocess.run(
            args,
            input=md_text,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=60,
        )
        if result.returncode == 0:
            return True
        # 出错时 stderr 会有 pandoc 的诊断信息
        print(f"[Pandoc] Error: {result.stderr[:500]}", file=sys.stderr)
        return False
    except FileNotFoundError:
        return False
    except subprocess.TimeoutExpired:
        print("[Pandoc] Timed out after 60s", file=sys.stderr)
        return False
    except Exception as e:
        print(f"[Pandoc] Unexpected error: {e}", file=sys.stderr)
        return False
