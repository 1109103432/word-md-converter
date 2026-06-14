# Word ↔ Markdown 转换工具

将 Word 文档(.docx) 和 Markdown(.md) 互相转换的桌面工具，**只需一个图标**，拖入文件即可自动识别转换方向。

## 功能特点

- 🖱️ **一键拖放** — 拖入 .docx 自动转 .md，拖入 .md 自动转 .docx，一个图标搞定
- 📑 **标题映射** — Word 标题样式（标题 1~6）与 Markdown 标题（#~######）完美对应
- 📝 **格式保留** — 保留粗体、斜体、删除线、代码块、超链接
- 📊 **表格转换** — 双向转换 Word 表格 ↔ Markdown 表格
- 🖼️ **图片提取** — Word 转 MD 时自动提取图片到本地
- 📋 **列表支持** — 有序列表和无序列表双向保留
- ⚙️ **参数设置** — 独立设置窗口，配置标题映射、通知参数等
- 🔔 **浮窗通知** — 转换完成后右下角弹出结果通知

## 安装方法

1. **确保已安装 Python 3.9+**
   - 下载：https://www.python.org/downloads/
   - 安装时勾选 "Add Python to PATH"

2. **运行安装脚本**
   - 双击 `install.bat`
   - 脚本会自动安装依赖并在桌面创建 3 个图标

## 使用方式

### 基本使用

将文件拖放到桌面图标上：

| 图标 | 功能 | 拖入文件类型 |
|------|------|-------------|
| **转换工具** | 智能双向转换（自动识别方向） | .docx, .md, .markdown |

转换后的文件将保存在原文件所在目录。

### 参数设置

双击桌面上的 **⚙ 转换设置** 图标，可以配置：
- **标题映射** — 自定义 Word 样式与 Markdown 标题的对应关系
- **图片处理** — 是否提取图片、图片文件夹名称
- **表格/列表/链接** — 是否保留对应格式
- **通知设置** — 浮窗自动关闭时间、是否显示输出路径

### 命令行使用

```bash
# 智能转换（自动识别方向）
python src/converter_launcher.py document.docx   # → .md
python src/converter_launcher.py document.md     # → .docx

# 或者直接使用方向明确的脚本
python src/word_to_md.py document.docx
python src/md_to_word.py document.md

# 设置界面
python src/settings_app.py
```

## 标题映射说明

### Word → Markdown

| Word 样式 | Markdown |
|-----------|----------|
| Heading 1 / 标题 1 | `#` |
| Heading 2 / 标题 2 | `##` |
| Heading 3 / 标题 3 | `###` |
| Heading 4 / 标题 4 | `####` |
| Heading 5 / 标题 5 | `#####` |
| Heading 6 / 标题 6 | `######` |

### Markdown → Word

| Markdown | Word 样式 |
|----------|----------|
| `#` | Heading 1 |
| `##` | Heading 2 |
| `###` | Heading 3 |
| `####` | Heading 4 |
| `#####` | Heading 5 |
| `######` | Heading 6 |

## 项目结构

```
doc-md-converter/
├── src/
│   ├── converter.py            # 核心转换引擎
│   ├── converter_launcher.py   # 统一智能转换入口（自动识别方向）
│   ├── word_to_md.py           # Word → MD 入口
│   ├── md_to_word.py           # MD → Word 入口
│   ├── settings_app.py         # 参数设置 GUI
│   ├── notification.py         # 浮窗通知模块
│   ├── config.py               # 配置管理
│   └── generate_icons.py       # 图标生成
├── icons/                 # 图标文件夹
├── config.json            # 用户配置文件
├── install.bat            # 安装脚本
├── requirements.txt       # 依赖列表
└── README.md              # 本文件
```

## 配置文件

所有设置保存在 `config.json` 中，可以手动编辑：

```json
{
  "heading_mapping": {
    "word_to_md": { "标题 1": "#", ... },
    "md_to_word": { "#": "Heading 1", ... }
  },
  "extract_images": true,
  "image_folder": "images",
  "preserve_tables": true,
  "output_encoding": "utf-8",
  "notification": {
    "auto_close_seconds": 5,
    "show_output_path": true
  }
}
```

## 依赖

- Python 3.9+
- [python-docx](https://python-docx.readthedocs.io/) — Word 文档读写
- [markdown](https://python-markdown.github.io/) — Markdown 解析
- tkinter — GUI（Python 自带）

## 许可

MIT
