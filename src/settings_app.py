"""
参数设置 GUI — 配置 Word ↔ Markdown 转换的各项参数。
可作为独立应用运行，或通过桌面"设置"图标启动。

标签页结构:
  1. MD→Word   — Pandoc 状态、样式模板、标题映射
  2. Word→MD   — 大纲级别识别、内容提取、图片、输出格式
  3. 其他设置   — 通知开关、时长、预览

支持 PyInstaller 打包和直接源码运行两种模式。
"""
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

# 源码运行时确保能找到同目录下的模块
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import load_config, save_config, get_config_path


class SettingsApp:
    """转换参数设置主窗口。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Word ↔ Markdown 转换设置")
        self.root.geometry("640x540")
        self.root.resizable(True, True)
        self.root.minsize(520, 420)

        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass

        self.config = load_config()
        self._setup_styles()
        self._build_ui()

        # 居中显示
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"+{x}+{y}")

    # ── 样式 ──────────────────────────────────────────────

    def _setup_styles(self):
        style = ttk.Style()
        # 使用 Windows 原生主题，勾选框显示 ✔ 而非 ❌
        try:
            style.theme_use("vista")
        except Exception:
            try:
                style.theme_use("winnative")
            except Exception:
                style.theme_use("default")

        style.configure("Title.TLabel", font=("Microsoft YaHei", 14, "bold"))
        style.configure("Section.TLabel", font=("Microsoft YaHei", 11, "bold"))
        style.configure("Desc.TLabel", font=("Microsoft YaHei", 9), foreground="#666")
        style.configure("Mono.TLabel", font=("Consolas", 10))
        style.configure("Save.TButton", font=("Microsoft YaHei", 10, "bold"), padding=8)
        style.configure("Future.TLabel", font=("Microsoft YaHei", 9),
                        foreground="#aaa", padding=(0, 8))

    # ── 主布局 ────────────────────────────────────────────

    def _build_ui(self):
        # 顶部标题栏
        title_frame = ttk.Frame(self.root, padding=(16, 12))
        title_frame.pack(fill=tk.X)

        ttk.Label(
            title_frame, text="⚙️  转换参数设置",
            style="Title.TLabel",
        ).pack(side=tk.LEFT)

        ttk.Label(
            title_frame,
            text=f"配置文件：{get_config_path()}",
            style="Desc.TLabel",
        ).pack(side=tk.RIGHT)

        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X)

        # Notebook
        notebook = ttk.Notebook(self.root, padding=(8, 4))
        notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # 三个标签页
        tab_md2word = ttk.Frame(notebook)
        notebook.add(tab_md2word, text="  MD → Word  ")
        self._build_md2word_tab(tab_md2word)

        tab_word2md = ttk.Frame(notebook)
        notebook.add(tab_word2md, text="  Word → MD  ")
        self._build_word2md_tab(tab_word2md)

        tab_other = ttk.Frame(notebook)
        notebook.add(tab_other, text="  其他设置  ")
        self._build_other_tab(tab_other)

        # 底部按钮
        btn_frame = ttk.Frame(self.root, padding=(12, 8))
        btn_frame.pack(fill=tk.X)

        ttk.Button(
            btn_frame, text="恢复默认设置",
            command=self._reset_defaults,
        ).pack(side=tk.LEFT, padx=4)

        ttk.Button(
            btn_frame, text="取消",
            command=self.root.destroy,
        ).pack(side=tk.RIGHT, padx=4)

        ttk.Button(
            btn_frame, text="💾  保存设置",
            style="Save.TButton",
            command=self._save_settings,
        ).pack(side=tk.RIGHT, padx=8)

    # ═══════════════════════════════════════════════════════
    #  Tab 1: MD → Word
    # ═══════════════════════════════════════════════════════

    def _build_md2word_tab(self, parent):
        frame = ttk.Frame(parent, padding=(16, 12))
        frame.pack(fill=tk.BOTH, expand=True)

        # ── Pandoc 状态 ──
        pandoc_frame = ttk.LabelFrame(frame, text="Pandoc 引擎", padding=12)
        pandoc_frame.pack(fill=tk.X, pady=(0, 12))

        from pandoc_engine import has_pandoc, get_pandoc_version
        if has_pandoc():
            ver = get_pandoc_version() or "未知"
            status = f"✅ Pandoc {ver} — 已就绪"
            desc = "所有 MD→Word 转换均通过 Pandoc 引擎处理（Typora 同款引擎）。"
        else:
            status = "⚠️ 未检测到 Pandoc"
            desc = "MD→Word 将回退到内置引擎，输出质量会降低。\n请将 pandoc.exe 放入应用目录。"

        ttk.Label(pandoc_frame, text=status, style="Section.TLabel").pack(
            anchor=tk.W, pady=(0, 4))
        ttk.Label(pandoc_frame, text=desc, style="Desc.TLabel").pack(anchor=tk.W)

        # ── 样式模板 ──
        template_frame = ttk.LabelFrame(
            frame,
            text="📄  输出样式模板（控制 docx 的字体、颜色、行距等外观）",
            padding=10,
        )
        template_frame.pack(fill=tk.X, pady=(0, 12))

        from pandoc_engine import get_builtin_template_path
        builtin_path = get_builtin_template_path()
        self._builtin_template_exists = builtin_path is not None

        self._template_labels = {
            "built-in": "内置模板 (template.docx)",
            "custom": "自定义模板...",
            "none": "不使用模板",
        }

        ref_value = self.config.get("pandoc_reference_doc", "built-in")
        self._custom_template_path = ""
        if ref_value == "built-in":
            self.template_mode = tk.StringVar(value=self._template_labels["built-in"])
        elif ref_value and ref_value.strip():
            self.template_mode = tk.StringVar(value=self._template_labels["custom"])
            self._custom_template_path = ref_value
        else:
            self.template_mode = tk.StringVar(value=self._template_labels["none"])

        mode_row = ttk.Frame(template_frame)
        mode_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(mode_row, text="模板选择：").pack(side=tk.LEFT, padx=(0, 8))

        self.template_combo = ttk.Combobox(
            mode_row,
            textvariable=self.template_mode,
            values=list(self._template_labels.values()),
            state="readonly",
            width=30,
        )
        self.template_combo.pack(side=tk.LEFT, padx=4)
        self.template_combo.bind("<<ComboboxSelected>>", self._on_template_mode_change)

        # 内置模板状态
        if self._builtin_template_exists:
            ttk.Label(
                template_frame,
                text=f"内置模板：{builtin_path}",
                style="Desc.TLabel",
            ).pack(anchor=tk.W, pady=(2, 0))
        else:
            ttk.Label(
                template_frame,
                text="⚠️ 未找到内置 template.docx，请将其放入应用目录",
                style="Desc.TLabel",
                foreground="#c00",
            ).pack(anchor=tk.W, pady=(2, 0))

        # 自定义路径（默认隐藏）
        self.custom_frame = ttk.Frame(template_frame)
        if self._get_template_mode_internal() == "custom":
            self.custom_frame.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(self.custom_frame, text="模板路径：").pack(side=tk.LEFT, padx=(0, 8))
        self.ref_doc_var = tk.StringVar(value=self._custom_template_path)
        ttk.Entry(self.custom_frame, textvariable=self.ref_doc_var, width=38).pack(
            side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        ttk.Button(self.custom_frame, text="浏览...",
                   command=self._browse_template).pack(side=tk.LEFT, padx=4)

        # ── 标题映射（固定，只读展示）──
        heading_frame = ttk.LabelFrame(
            frame,
            text="Markdown 标题 → Word 样式（固定映射，由 Pandoc 自动处理）",
            padding=10,
        )
        heading_frame.pack(fill=tk.X, pady=(0, 12))

        items = [
            ("#", "Heading 1"), ("##", "Heading 2"),
            ("###", "Heading 3"), ("####", "Heading 4"),
            ("#####", "Heading 5"), ("######", "Heading 6"),
        ]
        for md_pfx, word_style in items:
            row = ttk.Frame(heading_frame)
            row.pack(side=tk.LEFT, padx=(0, 16), pady=2)
            ttk.Label(row, text=f"{md_pfx}  →  {word_style}",
                      style="Desc.TLabel").pack()

        # ── 未来扩展 ──
        ttk.Label(
            frame,
            text="🔮  未来可扩展：页面设置（A4/Letter）、目录生成、"
                 "代码块高亮样式、图片分辨率控制",
            style="Future.TLabel",
        ).pack(anchor=tk.W)

    # ═══════════════════════════════════════════════════════
    #  Tab 2: Word → MD
    # ═══════════════════════════════════════════════════════

    def _build_word2md_tab(self, parent):
        frame = ttk.Frame(parent, padding=(16, 12))
        frame.pack(fill=tk.BOTH, expand=True)

        # ── 标题识别 ──
        heading_frame = ttk.LabelFrame(
            frame, text="🔍  标题识别（基于大纲级别）", padding=10)
        heading_frame.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(
            heading_frame,
            text="自动读取 Word 样式的大纲级别 (outlineLvl)，与样式名称无关：",
            style="Desc.TLabel",
        ).pack(anchor=tk.W, pady=(0, 6))

        levels = [
            ("Lv0", "#"), ("Lv1", "##"), ("Lv2", "###"),
            ("Lv3", "####"), ("Lv4", "#####"), ("Lv5", "######"),
        ]
        row = ttk.Frame(heading_frame)
        row.pack(fill=tk.X, pady=(0, 4))
        for lv, md in levels:
            ttk.Label(row, text=f"{lv} → {md}  ",
                      style="Mono.TLabel").pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(
            heading_frame,
            text="不同样式名但相同大纲级别 → 输出同一 MD 标题层级。"
                 "无需手动配置任何映射。",
            style="Desc.TLabel",
        ).pack(anchor=tk.W)

        # ── 内容提取 ──
        content_frame = ttk.LabelFrame(
            frame, text="内容提取", padding=10)
        content_frame.pack(fill=tk.X, pady=(0, 12))

        self.preserve_tables_var = tk.BooleanVar(
            value=self.config.get("preserve_tables", True))
        self.preserve_lists_var = tk.BooleanVar(
            value=self.config.get("preserve_lists", True))
        self.preserve_links_var = tk.BooleanVar(
            value=self.config.get("preserve_hyperlinks", True))

        cb_row = ttk.Frame(content_frame)
        cb_row.pack(fill=tk.X)
        ttk.Checkbutton(cb_row, text="转换表格",
                        variable=self.preserve_tables_var).pack(
            side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(cb_row, text="转换列表",
                        variable=self.preserve_lists_var).pack(
            side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(cb_row, text="保留超链接",
                        variable=self.preserve_links_var).pack(side=tk.LEFT)

        # ── 图片 ──
        img_frame = ttk.LabelFrame(frame, text="图片处理", padding=10)
        img_frame.pack(fill=tk.X, pady=(0, 12))

        self.extract_images_var = tk.BooleanVar(
            value=self.config.get("extract_images", True))
        ttk.Checkbutton(
            img_frame, text="提取 Word 文档中嵌入的图片",
            variable=self.extract_images_var,
        ).pack(anchor=tk.W, pady=(0, 4))

        img_sub = ttk.Frame(img_frame)
        img_sub.pack(fill=tk.X, padx=(24, 0))
        ttk.Label(img_sub, text="保存到文件夹：").pack(side=tk.LEFT)
        self.image_folder_var = tk.StringVar(
            value=self.config.get("image_folder", "images"))
        ttk.Entry(img_sub, textvariable=self.image_folder_var, width=18).pack(
            side=tk.LEFT, padx=8)

        # ── 输出格式 ──
        output_frame = ttk.LabelFrame(frame, text="输出格式", padding=10)
        output_frame.pack(fill=tk.X, pady=(0, 12))

        self.add_toc_var = tk.BooleanVar(
            value=self.config.get("add_toc", True))
        ttk.Checkbutton(
            output_frame,
            text="在 Markdown 开头添加 [TOC] 目录占位符",
            variable=self.add_toc_var,
        ).pack(anchor=tk.W, pady=(0, 6))

        enc_row = ttk.Frame(output_frame)
        enc_row.pack(fill=tk.X)
        ttk.Label(enc_row, text="输出编码：").pack(side=tk.LEFT)
        self.encoding_var = tk.StringVar(
            value=self.config.get("output_encoding", "utf-8"))
        ttk.Combobox(
            enc_row, textvariable=self.encoding_var,
            values=["utf-8", "utf-8-sig", "gbk", "gb2312"],
            width=12, state="readonly",
        ).pack(side=tk.LEFT, padx=8)

        # ── 未来扩展 ──
        ttk.Label(
            frame,
            text="🔮  未来可扩展：批注/修订提取、YAML 元数据头、"
                 "页面范围选择、图片格式转换",
            style="Future.TLabel",
        ).pack(anchor=tk.W)

    # ═══════════════════════════════════════════════════════
    #  Tab 3: 其他设置
    # ═══════════════════════════════════════════════════════

    def _build_other_tab(self, parent):
        frame = ttk.Frame(parent, padding=(16, 12))
        frame.pack(fill=tk.BOTH, expand=True)

        # ── 通知 ──
        notify_frame = ttk.LabelFrame(
            frame, text="🔔  Windows 系统通知", padding=12)
        notify_frame.pack(fill=tk.X, pady=(0, 12))

        self.notif_enabled_var = tk.BooleanVar(
            value=self.config.get("notification", {}).get("enabled", True))
        ttk.Checkbutton(
            notify_frame, text="启用通知（关闭后不显示任何通知）",
            variable=self.notif_enabled_var,
        ).pack(anchor=tk.W)

        self.show_success_var = tk.BooleanVar(
            value=self.config.get("notification", {}).get("show_success", True))
        ttk.Checkbutton(
            notify_frame, text="转换成功时也通知（关闭后仅在出错时弹出）",
            variable=self.show_success_var,
        ).pack(anchor=tk.W, pady=(4, 0))

        # 时长
        dur_row = ttk.Frame(notify_frame)
        dur_row.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(dur_row, text="显示时长：").pack(side=tk.LEFT)

        old_dur = self.config.get("notification", {}).get("duration", "5秒")
        self.duration_var = tk.StringVar(value=old_dur)
        ttk.Combobox(
            dur_row, textvariable=self.duration_var,
            values=["5秒", "10秒", "25秒", "持续", "紧急"],
            state="readonly",
            width=8,
        ).pack(side=tk.LEFT, padx=8)

        # 路径显示
        self.show_path_var = tk.BooleanVar(
            value=self.config.get("notification", {}).get("show_output_path", True))
        ttk.Checkbutton(
            notify_frame,
            text="显示输出文件路径（可点击打开所在文件夹）",
            variable=self.show_path_var,
        ).pack(anchor=tk.W, pady=(10, 0))

        # 预览
        preview_row = ttk.Frame(notify_frame)
        preview_row.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(
            preview_row, text="🔔  预览通知效果",
            command=self._preview_notification,
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(
            preview_row, text="查看当前设置下的通知效果",
            style="Desc.TLabel",
        ).pack(side=tk.LEFT)

        # ── 未来扩展 ──
        ttk.Label(
            frame,
            text="🔮  未来可扩展：界面语言、开机自启、文件关联、自动更新检查",
            style="Future.TLabel",
        ).pack(anchor=tk.W)

    # ── 模板选择逻辑 ──────────────────────────────────────

    def _get_template_mode_internal(self) -> str:
        """将 Combobox 显示标签转换回内部值。"""
        display = self.template_mode.get()
        for key, label in self._template_labels.items():
            if label == display:
                return key
        return "built-in"

    def _on_template_mode_change(self, event=None):
        if self._get_template_mode_internal() == "custom":
            self.custom_frame.pack(fill=tk.X, pady=(8, 0))
        else:
            self.custom_frame.pack_forget()

    def _browse_template(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="选择参考样式模板",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if path:
            self.ref_doc_var.set(path)
            self._custom_template_path = path

    # ── 通知预览 ──────────────────────────────────────────

    def _preview_notification(self):
        from notification import show_notification
        self.root.withdraw()
        show_notification(
            title="✅ 转换成功（预览）",
            message="这是通知预览效果。\n\n"
                    "源文件：example.docx\n"
                    "输出文件：example.md",
            output_path=str(Path.home() / "Desktop" / "example.md"),
            auto_close=self.duration_var.get(),
        )
        self.root.deiconify()

    # ── 保存 & 重置 ───────────────────────────────────────

    def _save_settings(self):
        # MD→Word 设置
        mode = self._get_template_mode_internal()
        if mode == "built-in":
            self.config["pandoc_reference_doc"] = "built-in"
        elif mode == "custom":
            self.config["pandoc_reference_doc"] = self.ref_doc_var.get()
        else:
            self.config["pandoc_reference_doc"] = ""

        # Word→MD 设置
        self.config["preserve_tables"] = self.preserve_tables_var.get()
        self.config["preserve_lists"] = self.preserve_lists_var.get()
        self.config["preserve_hyperlinks"] = self.preserve_links_var.get()
        self.config["extract_images"] = self.extract_images_var.get()
        self.config["image_folder"] = self.image_folder_var.get()
        self.config["add_toc"] = self.add_toc_var.get()
        self.config["output_encoding"] = self.encoding_var.get()

        # 通知设置
        self.config.setdefault("notification", {})["enabled"] = \
            self.notif_enabled_var.get()
        self.config.setdefault("notification", {})["show_success"] = \
            self.show_success_var.get()
        self.config.setdefault("notification", {})["duration"] = \
            self.duration_var.get()
        self.config.setdefault("notification", {})["show_output_path"] = \
            self.show_path_var.get()
        # 清理旧配置键
        self.config.setdefault("notification", {}).pop("auto_close_seconds", None)

        # 清理已不再使用的旧配置项
        self.config.pop("use_pandoc", None)
        heading = self.config.get("heading_mapping", {})
        heading.pop("word_to_md", None)

        if save_config(self.config):
            messagebox.showinfo("保存成功", "设置已保存到配置文件。")
        else:
            messagebox.showerror("保存失败", "无法写入配置文件，请检查文件权限。")

    def _reset_defaults(self):
        if messagebox.askyesno(
            "确认恢复默认",
            "确定要恢复所有设置为默认值吗？\n\n"
            "这将重置：\n"
            "  · 样式模板 → 内置模板\n"
            "  · 内容提取 → 全部开启\n"
            "  · 图片提取 → 开启\n"
            "  · 输出格式 → UTF-8，添加 TOC\n"
            "  · 通知 → 全部开启，25秒",
        ):
            from config import _DEFAULT_CONFIG
            _DEFAULT_CONFIG["notification"]["duration"] = "5秒"
            _DEFAULT_CONFIG["pandoc_reference_doc"] = "built-in"
            _DEFAULT_CONFIG["heading_mapping"] = {"md_to_word": {
                "#": "Heading 1", "##": "Heading 2", "###": "Heading 3",
                "####": "Heading 4", "#####": "Heading 5", "######": "Heading 6",
            }}
            save_config(_DEFAULT_CONFIG)
            self.root.destroy()
            new_root = tk.Tk()
            SettingsApp(new_root)
            new_root.mainloop()


def main():
    root = tk.Tk()
    SettingsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
