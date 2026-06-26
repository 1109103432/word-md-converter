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
    app_dir = _get_app_dir()

    # 2. 按优先级搜索
    search_dirs = [
        app_dir / "bin",                # bin/ 子目录（分发版首选）
        app_dir,                        # 应用根目录
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
    """获取应用根目录（与 config._get_app_dir 一致的逻辑）。

    分发版中 exe 位于 bin/ 子目录，因此需返回 bin/ 的父级。
    """
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "bin":
            return exe_dir.parent
        return exe_dir
    else:
        return Path(__file__).resolve().parent.parent


def get_builtin_template_path() -> Path | None:
    """返回内置模板 内置模板.docx 的路径，不存在则返回 None。

    仅在 模板/ 目录中查找。
    """
    template = _get_templates_dir() / "内置模板.docx"
    if template.is_file():
        return template
    return None


def _get_templates_dir() -> Path:
    """返回模板文件夹路径。"""
    return _get_app_dir() / "模板"


# ── 内置模板备份与恢复 ──

_BACKUP_NAME = ".内置模板.backup"


def _get_backup_path() -> Path:
    """返回内置模板备份文件的路径（隐藏文件，存放在应用根目录）。"""
    return _get_app_dir() / _BACKUP_NAME


def create_builtin_backup() -> bool:
    """
    为内置模板创建一份隐藏备份（.内置模板.backup），放在应用根目录。

    规则：
      - 备份已存在 → 跳过（不覆盖，保证是原始副本）
      - 内置模板存在 → 复制为备份
      - 内置模板不存在 → 无法创建，返回 False

    此备份用于"恢复默认设置"时将内置模板还原到原始状态。
    """
    backup = _get_backup_path()
    if backup.is_file():
        return True  # 已有备份，不覆盖

    source = get_builtin_template_path()
    if source and source.is_file():
        import shutil
        shutil.copy2(str(source), str(backup))
        return True

    return False


def restore_builtin_from_backup() -> bool:
    """
    用隐藏备份覆盖 模板/内置模板.docx。

    即使内置模板被用户删除或修改，此操作都能恢复为原始文件。
    """
    backup = _get_backup_path()
    if not backup.is_file():
        return False

    import shutil

    target_dir = _get_templates_dir()
    target = target_dir / "内置模板.docx"
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        if target.is_file():
            target.chmod(0o644)  # 解除只读，确保可覆盖
        shutil.copy2(str(backup), str(target))
        return True
    except OSError:
        return False


def resolve_reference_doc(value: str = None) -> str | None:
    """
    将配置中的 pandoc_reference_doc 值解析为实际文件路径。

    解析规则：
      - "built-in" 或空字符串 → 查找内置 内置模板.docx
      - "custom:<filename>" → 从 模板/ 目录查找
      - 其他路径值 → 作为文件路径，存在则使用

    返回路径字符串或 None（回退到无模板）。
    """
    if not value or value == "built-in":
        # 使用内置模板（应用根目录的 内置模板.docx）
        template = get_builtin_template_path()
        return str(template) if template else None

    if value.startswith("custom:"):
        # 自定义模板：从 模板/ 目录加载
        filename = value[len("custom:"):]
        template = _get_templates_dir() / filename
        if template.is_file():
            return str(template)
        # 文件已被删除 → 回退到内置模板
        builtin = get_builtin_template_path()
        return str(builtin) if builtin else None

    # 兼容旧格式：直接路径
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
                      支持 "built-in"（使用内置 内置模板.docx）、
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
