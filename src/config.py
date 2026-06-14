"""
配置管理模块 — 读取、写入、合并配置文件。
支持常规 Python 运行和 PyInstaller 打包后的 frozen 模式。
"""
import json
import sys
from pathlib import Path

_DEFAULT_CONFIG = {
    "heading_mapping": {
        "md_to_word": {
            "#": "Heading 1", "##": "Heading 2", "###": "Heading 3",
            "####": "Heading 4", "#####": "Heading 5", "######": "Heading 6",
        },
    },
    "extract_images": True,
    "image_folder": "images",
    "preserve_tables": True,
    "preserve_lists": True,
    "preserve_hyperlinks": True,
    "output_encoding": "utf-8",
    "add_toc": True,
    "pandoc_reference_doc": "built-in",
    "notification": {
        "enabled": True,
        "show_success": True,
        "duration": "5秒",
        "show_output_path": True,
    },
}


def _get_app_dir() -> Path:
    """获取应用根目录。

    在 PyInstaller frozen 模式下，返回 .exe 所在目录；
    在普通 Python 模式下，返回项目根目录（本文件向上两级）。
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后：exe 所在目录
        return Path(sys.executable).resolve().parent
    else:
        # 开发模式：本文件在 src/ 下，项目根目录在 src/ 的父级
        return Path(__file__).resolve().parent.parent


# 应用根目录和配置文件路径（延迟计算，以支持 frozen 模式）
_APP_DIR: Path | None = None
_CONFIG_PATH: Path | None = None


def _init_paths():
    """初始化应用目录和配置文件路径。"""
    global _APP_DIR, _CONFIG_PATH
    if _APP_DIR is None:
        _APP_DIR = _get_app_dir()
        _CONFIG_PATH = _APP_DIR / "config.json"


def _deep_merge(base: dict, override: dict) -> dict:
    """深度合并两个字典，override 覆盖 base。"""
    result = base.copy()
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_config() -> dict:
    """加载配置文件，如不存在则返回默认配置。"""
    _init_paths()
    if _CONFIG_PATH.exists():
        try:
            with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
                user_config = json.load(f)
            return _deep_merge(_DEFAULT_CONFIG, user_config)
        except (json.JSONDecodeError, IOError):
            pass
    return _DEFAULT_CONFIG.copy()


def save_config(config: dict) -> bool:
    """保存配置到 JSON 文件。"""
    _init_paths()
    try:
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except IOError:
        return False


def get_config_path() -> Path:
    """返回配置文件路径。"""
    _init_paths()
    return _CONFIG_PATH


def get_project_root() -> Path:
    """返回项目根目录。"""
    _init_paths()
    return _APP_DIR
