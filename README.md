# Word-MD快速转换

**一个图标的双向文档转换器** — 拖入文件自动识别方向，双击图标剪贴板秒变 Word。

![version](https://img.shields.io/badge/version-2.2.1-blue)
![platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)
![license](https://img.shields.io/badge/license-MIT-green)

---

## ✨ 功能特点

| 功能 | 说明 |
|------|------|
| 🔄 **智能双向转换** | 拖入 `.docx` 自动转 `.md`，拖入 `.md` 自动转 `.docx`，无需手动选择方向 |
| 📋 **剪贴板秒转 Word** | 双击图标，剪贴板里的 Markdown 内容直接变成 Word 文档（⭐ 核心巧思） |
| 📑 **标题完美映射** | Word 大纲级别 ⇄ Markdown `#`~`######`，层级精准保留 |
| 📊 **表格双向保留** | Word 表格 ↔ Markdown 表格，格式无损 |
| 🖼️ **图片自动提取** | Word→MD 时图片自动导出到本地文件夹 |
| 📝 **内联格式** | 粗体、斜体、删除线、行内代码、超链接 — 全部保留 |
| 📐 **Word 格式可控** | MD→Word 严格按模板样式导出，内置通用模板，支持自定义模板切换 |
| 🔔 **原生通知** | Windows 10/11 系统通知，点通知直接打开文件所在目录 |
| ⚙️ **可视化设置** | 模板选择、标题映射、开关项……无需编辑配置文件 |
| 📦 **免安装运行** | 解压即用，无需 Python 环境，无需管理员权限 |

---

## 🎯 核心巧思：剪贴板 → Word

**这是本工具最特别的功能。**

设想这个场景：你在网页、编辑器、聊天记录里看到一段 Markdown 格式的文字，想把它们整理成规整的 Word 文档。

传统做法：
1. 打开 Word
2. 粘贴（格式全乱）
3. 逐一调整标题、加粗、表格……
4. 保存

本工具的做法：
1. **复制** Markdown 内容（`Ctrl+C`）
2. **双击** 桌面「开始转换」图标
3. **完成** — 一个排版精美的 Word 文档已经出现在你面前

文件名自动取 Markdown 第一个标题，无标题则用时间戳。这就是「剪贴板直转」的设计初衷——让 Markdown 到 Word 和复制粘贴一样简单。

---

## 🧠 用户需求 & 解决场景

| 谁会用 | 什么场景 | 本工具解决什么 |
|--------|---------|---------------|
| 📝 **写作者** | 用 Markdown 写初稿，投稿需 Word 格式 | 一键转 Word，排版自动处理 |
| 🔧 **文字工作者** | AI 输出的 Markdown 文本，需转为格式严谨的 Word 文档 | 复制 → 双击图标 → 一键转为模板样式统一的 Word |
| 📚 **学生/研究者** | 收集的 Word 资料想归档为 Markdown | 自动提取图片、保留表格层级 |
| 📄 **办公人员** | Word 文档需发布到支持 Markdown 的平台 | Word→MD 保留所有格式细节 |
| 💬 **所有人** | 看到一段好内容想存为文档 | 复制 → 双击图标 → 得到 Word |

---

## 🔬 工作原理

```
┌─────────────────────────────────────────────────────┐
│                    用户输入                          │
│  拖入文件(.docx/.md)  or  双击图标(读剪贴板)          │
└─────────────────┬───────────────────────────────────┘
                  ▼
┌─────────────────────────────────────────────────────┐
│              自动识别转换方向                         │
│   .docx/.md/.markdown/.txt/.text → 自动判定           │
└─────────────────┬───────────────────────────────────┘
                  ▼
┌─────────────────────────────────────────────────────┐
│                Pandoc 转换引擎                        │
│  • Word→MD: lxml修复样式→Pandoc gfm格式转换           │
│  • MD→Word: Pandoc docx输出+reference-doc模板         │
│  • Pandoc不可用时回退到 python-docx 原生引擎            │
└─────────────────┬───────────────────────────────────┘
                  ▼
┌─────────────────────────────────────────────────────┐
│              输出 + Windows 原生通知                   │
│  转换文件保存到源文件目录 + 右下角 Toast 通知           │
│  点击通知 → 打开文件所在文件夹                         │
└─────────────────────────────────────────────────────┘
```

**技术栈**：Python + PySide6 (Qt GUI) + Pandoc + python-docx + lxml + win11toast

---

## ⚠️ 已知缺陷

| 问题 | 严重程度 | 说明 |
|------|---------|------|
| 🔧 **设置窗口打开闪烁** | 中等 | 打开设置窗口时会有短暂的白屏闪烁，已尝试多种 Qt 方案修复（alpha 透明、geometry 预设、withdraw/update_idletasks），未能彻底解决。**欢迎 PR 帮助** |
| 📦 **体积较大（~280 MB）** | 低 | Pandoc（221 MB）+ PySide6 Qt 框架（~50 MB），单 exe 合并已节省 ~56 MB。Pandoc 是 Haskell 编写的大型万能转换器，暂时无法替代 |
| 🖼️ **Word 复杂排版丢失** | 低 | 文本框、艺术字、页眉页脚、分栏、浮动图片等高级 Word 特性无法转为 Markdown |
| 📊 **合并单元格表格** | 低 | 原生引擎不支持合并单元格，Pandoc 引擎部分支持 |
| 🖥️ **仅支持 Windows** | 中 | 依赖 Windows Toast 通知和 win11toast，macOS/Linux 用户需修改通知模块 |
| 🔤 **.doc 有限支持** | 低 | `.doc` 通过 Pandoc 直接读取，图片/表格等复杂内容可能丢失，建议先用 Word 另存为 `.docx` |

---

## 🙏 关于作者

我是一名 **没有编程背景的普通用户**，这个工具是用 AI 辅助（Claude Code）一行一行"聊"出来的。

坦白说，我不会写代码。这个项目的每一行 Python 都是在 AI 的帮助下完成的——我描述需求，AI 生成代码，我测试反馈，AI 修复问题——如此反复数百轮。

正因为如此：
- 代码中可能存在不规范的地方，请多包涵
- 如果你发现 bug 或有改进建议，请提 [Issue](https://github.com/issues)
- 如果你擅长 Qt/PySide6 并知道如何修复**设置窗口打开闪烁**的问题，**非常需要你的帮助**

---

## 🆘 求助：设置窗口打开闪烁

### 问题描述

双击「转换设置」打开设置窗口时，窗口会短暂白屏闪烁一下才显示内容。

### 已尝试的方案（均未彻底解决）

1. `setAttribute(Qt.WA_TranslucentBackground)` + alpha 渐变 — 窗口变全透明
2. `geometry()` 一次性设置大小+位置 — 闪得更厉害
3. `withdraw()` + `update_idletasks()` + `show()` — 闪烁依旧

### 环境

- Python 3.9+ / PySide6 6.5+
- Qt Fusion 样式
- Windows 11

### 如果你知道解决方案

请在 [Issues](https://github.com/issues) 区留言或直接提 PR，非常非常感谢 🙏

---

## 📥 安装与使用

### 方式一：下载打包版（推荐，无需 Python）

1. 下载 `Word-MD快速转换-v2.2.1.zip`
2. 解压到任意目录（建议不放在桌面）
3. 双击 `安装.bat` → 桌面出现两个图标
4. 开始使用！

### 方式二：源码运行（需 Python 3.9+）

```bash
git clone https://github.com/1109103432/word-md-converter.git
cd doc-md
pip install -r requirements.txt
python setup_shortcuts.py   # 创建桌面快捷方式
```

---

## 📂 项目结构

```
doc-md/
├── src/
│   ├── launcher.py               # 统一启动入口 (--settings 切换模式)
│   ├── converter.py              # 核心转换引擎 (Word↔MD)
│   ├── converter_launcher.py     # 转换器逻辑 (拖放+剪贴板)
│   ├── settings_app.py           # 转换设置 GUI
│   ├── pandoc_engine.py          # Pandoc 引擎 + 模板管理
│   ├── notification.py           # Windows Toast 通知
│   └── config.py                 # 配置管理
├── 模板/
│   └── 内置模板.docx              # 内置 Word 参考模板
├── icons/                        # 图标文件
├── build.py                      # PyInstaller 打包脚本
├── config.json                   # 默认配置
├── requirements.txt              # Python 依赖
└── README.md                     # 本文件
```

---

## 🤝 贡献

欢迎任何形式的贡献：
- 🐛 报告 bug
- 💡 提出新功能建议
- 🔧 修复设置窗口闪烁问题（急需！）
- 📖 改进文档

---

## 📄 许可

MIT License — 可自由使用、修改、分发。
