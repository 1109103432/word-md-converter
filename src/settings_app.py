"""
参数设置 GUI — 配置 Word ↔ Markdown 转换的各项参数。
可作为独立应用运行，或通过桌面"设置"图标启动。

标签页结构:
  1. MD→Word   — Pandoc 状态、样式模板选择
  2. Word→MD   — 大纲级别识别、内容提取、图片、输出格式
  3. 其他设置   — 通知开关、时长、预览

技术栈: PySide6 (Qt for Python)，彻底解决 tkinter 窗口闪烁问题。
"""
import sys
import os
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QTabWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QLabel,
    QPushButton,
    QCheckBox,
    QComboBox,
    QLineEdit,
    QButtonGroup,
    QMessageBox,
    QFrame,
    QSizePolicy,
)
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QFont, QIcon, QPainter, QPixmap, QPen, QColor

# 源码运行时确保能找到同目录下的模块
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import load_config, save_config, get_config_path, get_project_root
from pandoc_engine import has_pandoc, get_pandoc_version, get_builtin_template_path, \
    create_builtin_backup, restore_builtin_from_backup


# ═══════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════

# 窗口
WINDOW_WIDTH = 760
WINDOW_HEIGHT = 650
WINDOW_MIN_WIDTH = 620
WINDOW_MIN_HEIGHT = 520

# 配色 (Flat Design Professional)
COLOR_PRIMARY = "#2563EB"
COLOR_PRIMARY_DARK = "#1D4ED8"
COLOR_PRIMARY_LIGHT = "#DBEAFE"
COLOR_SUCCESS = "#10B981"
COLOR_WARNING = "#F59E0B"
COLOR_ERROR = "#EF4444"
COLOR_BG_WHITE = "#FFFFFF"
COLOR_BG_SURFACE = "#F8FAFC"
COLOR_BORDER = "#E2E8F0"
COLOR_TEXT_PRIMARY = "#1E293B"
COLOR_TEXT_SECONDARY = "#64748B"
COLOR_TEXT_MUTED = "#94A3B8"


# ═══════════════════════════════════════════════════════════
#  辅助函数
# ═══════════════════════════════════════════════════════════

def _get_templates_dir() -> Path:
    """返回模板文件夹路径，不存在则自动创建。"""
    root = get_project_root()
    templates_dir = root / "模板"
    templates_dir.mkdir(parents=True, exist_ok=True)
    return templates_dir


def _scan_templates() -> list[str]:
    """
    扫描 模板/ 目录，返回自定义模板文件名列表（不含内置模板）。
    内置模板（内置模板.docx）始终排除在外，由内置模板独立管理。
    按文件名排序。
    """
    templates_dir = _get_templates_dir()
    builtin_path = get_builtin_template_path()

    custom = []
    for f in sorted(templates_dir.glob("*.docx"), key=lambda p: p.name.lower()):
        # 排除内置模板（如果在 模板 目录下有同名文件也排除）
        if builtin_path and f.resolve() == builtin_path.resolve():
            continue
        # 也排除名为 内置模板.docx 的文件（保护内置模板不受影响）
        if f.name.lower() == "内置模板.docx":
            continue
        custom.append(f.name)
    return custom


def _builtin_template_exists() -> bool:
    """内置模板是否存在。"""
    return get_builtin_template_path() is not None


def _parse_template_config(value: str) -> tuple[str, str | None]:
    """
    解析 config 中的 pandoc_reference_doc 值。
    Returns:
        (mode, filename)
        mode: "built-in" | "custom" | "missing"
        filename: 自定义模板的文件名（仅 custom 模式）
    """
    if not value or value == "built-in":
        return ("built-in", None)

    if value.startswith("custom:"):
        filename = value[len("custom:"):]
        # 验证文件是否存在
        template_path = _get_templates_dir() / filename
        if template_path.is_file():
            return ("custom", filename)
        else:
            return ("missing", filename)

    # 兼容旧格式：直接路径
    path = Path(value)
    if path.is_file():
        return ("custom", path.name)

    return ("missing", None)


# ═══════════════════════════════════════════════════════════
#  样式表
# ═══════════════════════════════════════════════════════════

STYLESHEET = """
/* ── 全局 ── */
QMainWindow {
    background-color: #FFFFFF;
}
QWidget {
    font-family: "Segoe UI", "Microsoft YaHei", sans-serif;
    font-size: 14px;
    color: #1E293B;
}

/* ── 标签页 ── */
QTabWidget::pane {
    border: 1px solid #E2E8F0;
    background-color: #FFFFFF;
    border-radius: 4px;
    top: -1px;
}
QTabBar::tab {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    padding: 11px 22px;
    margin-right: 2px;
    border-radius: 4px 4px 0 0;
    color: #64748B;
    font-size: 14px;
}
QTabBar::tab:selected {
    background: #FFFFFF;
    color: #2563EB;
    border-bottom: 2px solid #2563EB;
    font-weight: 600;
}
QTabBar::tab:hover:!selected {
    color: #1E293B;
    background: #F1F5F9;
}

/* ── 分组框 ── */
QGroupBox {
    font-weight: 600;
    border: 1px solid #E2E8F0;
    border-radius: 6px;
    margin-top: 8px;
    padding-top: 20px;
    background-color: #FFFFFF;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 8px;
    color: #1E293B;
}

/* ── 按钮 ── */
QPushButton {
    border: 1px solid #E2E8F0;
    border-radius: 5px;
    padding: 9px 18px;
    background-color: #FFFFFF;
    color: #1E293B;
    font-size: 14px;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #CBD5E1;
}
QPushButton:pressed {
    background-color: #E2E8F0;
}
QPushButton#saveBtn {
    background-color: #2563EB;
    color: #FFFFFF;
    border: none;
    padding: 9px 26px;
    font-weight: 600;
}
QPushButton#saveBtn:hover {
    background-color: #1D4ED8;
}
QPushButton#saveBtn:pressed {
    background-color: #1E40AF;
}
QPushButton#resetBtn {
    border: none;
    color: #64748B;
    background: transparent;
    padding: 9px 18px;
}
QPushButton#resetBtn:hover {
    color: #EF4444;
    background: transparent;
}

/* ── 输入框 ── */
/*
 * 注意：QSS 的 vertical padding 会与 Fusion 风格的内部 content-area
 * 计算冲突，导致文字区域被挤扁 → 文字裁切。
 * 这里只用水平 padding 给文字呼吸空间，垂直方向通过 min-height
 * 交给 Fusion 风格处理，避免冲突。
 */
QLineEdit {
    border: 1px solid #E2E8F0;
    border-radius: 4px;
    padding: 0 10px;
    background: #FFFFFF;
    font-size: 14px;
    min-height: 30px;
}
QLineEdit:focus {
    border-color: #2563EB;
}

/* ── 下拉框 ── */
QComboBox {
    border: 1px solid #E2E8F0;
    border-radius: 4px;
    padding: 0 10px;
    background: #FFFFFF;
    min-width: 100px;
    font-size: 14px;
    min-height: 30px;
}
QComboBox:hover {
    border-color: #CBD5E1;
}
QComboBox:focus {
    border-color: #2563EB;
}
QComboBox::drop-down {
    border: none;
    width: 26px;
}
QComboBox QAbstractItemView {
    border: 1px solid #E2E8F0;
    background: #FFFFFF;
    selection-background-color: #DBEAFE;
    selection-color: #1E293B;
    padding: 6px;
}

/* ── 复选框（统一方框 ☑️ 样式）── */
QCheckBox {
    spacing: 10px;
    padding: 6px 0;
}
QCheckBox::indicator {
    width: 20px;
    height: 20px;
    border-radius: 4px;
}
QCheckBox::indicator:unchecked {
    border: 2px solid #CBD5E1;
    background: #FFFFFF;
}
QCheckBox::indicator:checked {
    border: 2px solid #2563EB;
    background: #2563EB;
    border-radius: 4px;
    image: url(CHECKMARK_PLACEHOLDER);
}

/* ── 分隔线 ── */
QFrame#separator {
    background-color: #E2E8F0;
    max-height: 1px;
    min-height: 1px;
}

/* ── 滚动区域 ── */
QScrollArea {
    border: none;
    background: transparent;
}

/* ── 提示/说明文字 ── */
QLabel#descLabel {
    color: #64748B;
    font-size: 13px;
}
QLabel#successLabel {
    color: #10B981;
    font-weight: 600;
}
QLabel#warningLabel {
    color: #F59E0B;
    font-weight: 600;
}
QLabel#sectionTitle {
    font-size: 15px;
    font-weight: 600;
    color: #1E293B;
}
QLabel#codeLabel {
    font-family: "Cascadia Code", "Consolas", "Courier New", monospace;
    font-size: 13px;
    color: #64748B;
}
"""


# ═══════════════════════════════════════════════════════════
#  勾选图标 — 用 QPainter 绘制 ✔ 保存为 PNG，QSS 中引用
# ═══════════════════════════════════════════════════════════

def _ensure_checkmark_icon() -> Path:
    """生成 checkmark 图标，返回文件路径。只生成一次，后续直接复用。"""
    icon_path = get_project_root() / "checkmark.png"
    if icon_path.exists():
        return icon_path

    # 4x 分辨率绘制 → 缩放到 20px 指示器时边缘锐利无模糊
    size = 80
    pixmap = QPixmap(size, size)
    pixmap.fill(QColor(0, 0, 0, 0))

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)

    # ✔ 折线 (相对于 80×80，等比缩放)
    pen = QPen(QColor("#FFFFFF"), 7.0, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
    painter.setPen(pen)
    painter.drawPolyline([
        QPoint(14, 46),
        QPoint(32, 62),
        QPoint(66, 18),
    ])

    painter.end()
    pixmap.save(str(icon_path), "PNG")
    return icon_path


# ═══════════════════════════════════════════════════════════
#  主窗口
# ═══════════════════════════════════════════════════════════

class SettingsWindow(QMainWindow):
    """转换参数设置主窗口 (PySide6)。"""

    def __init__(self):
        super().__init__()

        # ── 加载配置 ──
        self.config = load_config()

        # ── 模板扫描 ──
        self._custom_templates: list[str] = _scan_templates()
        self._builtin_available = _builtin_template_exists()

        # ── 修改追踪 ──
        self._dirty = False

        # ── 配置窗口 ──
        self._setup_window()
        self._build_ui()
        self._load_config_values()
        self._setup_dirty_tracking()
        self._center_on_screen()

    # ── 窗口设置 ──────────────────────────────────────────

    def _setup_window(self):
        """配置窗口属性。"""
        self.setWindowTitle("Word ↔ Markdown 转换设置")
        self.resize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.setMinimumSize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)

        # 窗口图标（尝试加载）
        try:
            icon_path = get_project_root() / "icons" / "settings.ico"
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass

    def _center_on_screen(self):
        """将窗口居中于屏幕。"""
        screen = QApplication.primaryScreen()
        if screen:
            center = screen.availableGeometry().center()
            geo = self.frameGeometry()
            geo.moveCenter(center)
            self.move(geo.topLeft())

    # ── 关闭拦截 ──────────────────────────────────────────

    def closeEvent(self, event):
        """关闭窗口前检查是否有未保存的修改。"""
        if self._dirty:
            msg = QMessageBox(self)
            msg.setWindowTitle("未保存的修改")
            msg.setText("设置已修改，是否保存？")
            msg.setIcon(QMessageBox.Question)
            save_btn = msg.addButton("保存", QMessageBox.AcceptRole)
            discard_btn = msg.addButton("不保存", QMessageBox.DestructiveRole)
            cancel_btn = msg.addButton("取消", QMessageBox.RejectRole)
            msg.setDefaultButton(save_btn)
            msg.exec()

            clicked = msg.clickedButton()
            if clicked == save_btn:
                self._save_settings()
                event.accept()
            elif clicked == discard_btn:
                event.accept()
            else:  # cancel_btn or closed via X
                event.ignore()
        else:
            event.accept()

    def _mark_dirty(self):
        """标记配置已被修改。"""
        self._dirty = True

    def _setup_dirty_tracking(self):
        """为所有可修改的控件绑定变更信号，实现未保存提醒。"""
        # QCheckBox — toggled
        checkboxes = [
            self.builtin_check,
            self.preserve_tables_cb,
            self.preserve_lists_cb,
            self.preserve_links_cb,
            self.extract_images_cb,
            self.add_toc_cb,
            self.notif_enabled_cb,
            self.show_success_cb,
            self.show_path_cb,
        ]
        for cb in checkboxes:
            cb.toggled.connect(self._mark_dirty)
        for cb in self.custom_checks:
            cb.toggled.connect(self._mark_dirty)

        # QLineEdit — textChanged
        self.image_folder_input.textChanged.connect(self._mark_dirty)

        # QComboBox — currentIndexChanged
        self.encoding_combo.currentIndexChanged.connect(self._mark_dirty)
        self.duration_combo.currentIndexChanged.connect(self._mark_dirty)

    # ── 文字控件辅助 ──────────────────────────────────────

    @staticmethod
    def _make_label(
        text: str = "",
        *,
        word_wrap: bool = False,
        object_name: str = "",
        bold: bool = False,
        font_size: int = 0,
        font_family: str = "",
        min_width: int = 0,
    ) -> QLabel:
        """
        统一创建 QLabel，自动设置 Preferred 尺寸策略防止文字被压缩裁切。

        规则:
          - 所有 label 默认 QSizePolicy.Preferred（垂直方向不压缩）
          - word_wrap=True 时水平方向也设为 Preferred
          - 后续文字不会因 layout 空间不足而被切掉
        """
        label = QLabel(text)
        label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)

        if word_wrap:
            label.setWordWrap(True)

        if object_name:
            label.setObjectName(object_name)

        if bold or font_size or font_family:
            font = label.font()
            if bold:
                font.setBold(True)
            if font_size:
                font.setPointSize(font_size)
            if font_family:
                font.setFamily(font_family)
            label.setFont(font)

        if min_width:
            label.setMinimumWidth(min_width)

        return label

    @staticmethod
    def _protect_text_widget(widget):
        """
        为输入控件设置正确的尺寸策略，防止 layout 压缩导致文字裁切。

        规则:
          - QLineEdit: Expanding 水平 + Preferred 垂直 + minHeight 28px
          - QComboBox: Preferred 水平 + Preferred 垂直 + minHeight 28px
          - minHeight 兜底: 即使 QSS 不生效，代码层也保证最低高度
        """
        if isinstance(widget, QLineEdit):
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        else:
            widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        widget.setMinimumHeight(30)
        return widget

    # ── 主布局 ────────────────────────────────────────────

    def _build_ui(self):
        """构建完整的 UI 结构。"""
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ── 顶部标题栏 ──
        main_layout.addWidget(self._build_header())

        # ── 分隔线 ──
        sep = QFrame()
        sep.setObjectName("separator")
        main_layout.addWidget(sep)

        # ── 标签页 ──
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_md2word_tab(), "MD → Word")
        self.tabs.addTab(self._build_word2md_tab(), "Word → MD")
        self.tabs.addTab(self._build_other_tab(), "其他设置")
        main_layout.addWidget(self.tabs, 1)  # stretch=1, 填充剩余空间

        # ── 底部按钮栏 ──
        main_layout.addWidget(self._build_footer())

    def _build_header(self) -> QWidget:
        """顶部标题栏：标题 + 配置文件路径。"""
        header = QWidget()
        header.setFixedHeight(56)
        layout = QHBoxLayout(header)
        layout.setContentsMargins(18, 14, 18, 14)

        title = self._make_label("⚙ 转换参数设置",
                                 object_name="sectionTitle", bold=True, font_size=15)
        layout.addWidget(title)

        layout.addStretch()

        config_label = self._make_label(f"配置文件：{get_config_path()}",
                                        object_name="descLabel")
        layout.addWidget(config_label)

        return header

    def _build_footer(self) -> QWidget:
        """底部按钮栏：恢复默认 / 取消 / 保存。"""
        footer = QWidget()
        footer.setFixedHeight(60)
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(14, 12, 14, 12)

        reset_btn = QPushButton("恢复默认设置")
        reset_btn.setObjectName("resetBtn")
        reset_btn.clicked.connect(self._reset_defaults)
        layout.addWidget(reset_btn)

        layout.addStretch()

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.close)
        layout.addWidget(cancel_btn)

        save_btn = QPushButton("保存设置")
        save_btn.setObjectName("saveBtn")
        save_btn.clicked.connect(self._save_settings)
        layout.addWidget(save_btn)

        return footer

    # ═══════════════════════════════════════════════════════
    #  Tab 1: MD → Word
    # ═══════════════════════════════════════════════════════

    def _build_md2word_tab(self) -> QWidget:
        """构建 MD→Word 标签页。"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(12)

        # ── Pandoc 状态 ──
        layout.addWidget(self._build_pandoc_status())

        # ── 样式模板 ──
        layout.addWidget(self._build_template_section())

        # ── 标题映射（只读展示）──
        layout.addWidget(self._build_heading_mapping())

        layout.addStretch()

        return tab

    def _build_pandoc_status(self) -> QWidget:
        """Pandoc 引擎状态显示。"""
        group = QGroupBox("Pandoc 引擎")
        layout = QVBoxLayout(group)
        layout.setSpacing(4)

        if has_pandoc():
            ver = get_pandoc_version() or "未知"
            status = self._make_label(f"✅ Pandoc {ver} — 已就绪",
                                      object_name="successLabel")
            desc = self._make_label(
                "所有 MD→Word 转换均通过 Pandoc 引擎处理（Typora 同款引擎）。",
                object_name="descLabel")
        else:
            status = self._make_label("⚠️ 未检测到 Pandoc",
                                      object_name="warningLabel")
            desc = self._make_label(
                "MD→Word 将回退到内置引擎，输出质量会降低。\n"
                "请将 pandoc.exe 放入应用目录。",
                object_name="descLabel")

        layout.addWidget(status)
        layout.addWidget(desc)

        return group

    def _build_template_section(self) -> QWidget:
        """样式模板选择区域。"""
        group = QGroupBox("📄 输出样式模板（控制 docx 的字体、颜色、行距等外观）")
        layout = QVBoxLayout(group)
        layout.setSpacing(10)

        # ── 两个操作按钮 ──
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        open_builtin_btn = QPushButton("打开内置模板")
        open_builtin_btn.setToolTip(
            str(get_builtin_template_path()) if self._builtin_available
            else "内置模板不存在"
        )
        open_builtin_btn.clicked.connect(self._open_builtin_template)
        if not self._builtin_available:
            open_builtin_btn.setEnabled(False)
        btn_row.addWidget(open_builtin_btn)

        open_folder_btn = QPushButton("打开模板文件夹")
        open_folder_btn.setToolTip(str(_get_templates_dir()))
        open_folder_btn.clicked.connect(self._open_templates_folder)
        btn_row.addWidget(open_folder_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        # ── 提示文字 ──
        hint = self._make_label(
            "💡 将 .docx 模板文件放入模板文件夹重新打开软件即可识别，"
            "可修改内置模板并另存为新模板。",
            object_name="descLabel", word_wrap=True)
        layout.addWidget(hint)

        # ── 模板选择（单选列表）──
        select_label = self._make_label("模板选择：", bold=True)
        layout.addWidget(select_label)

        # 复选框容器（QButtonGroup.setExclusive 保证互斥，行为同单选）
        self.template_check_group = QButtonGroup(self)
        self.template_check_group.setExclusive(True)

        check_container = QWidget()
        check_layout = QVBoxLayout(check_container)
        check_layout.setContentsMargins(8, 2, 0, 2)
        check_layout.setSpacing(6)

        # 内置模板 — 始终存在、默认选中、不可删除
        self.builtin_check = QCheckBox("内置模板.docx（默认）")
        self.builtin_check.setChecked(True)
        self.builtin_check.toggled.connect(self._on_template_changed)
        self.builtin_template_path = get_builtin_template_path()
        if self._builtin_available:
            self.builtin_check.setToolTip(str(self.builtin_template_path))
        else:
            self.builtin_check.setEnabled(False)
            self.builtin_check.setText("内置模板.docx（未找到）")
        check_layout.addWidget(self.builtin_check)
        self.template_check_group.addButton(self.builtin_check, id=0)

        # 自定义模板 — 动态扫描
        self.custom_checks: list[QCheckBox] = []
        for i, filename in enumerate(self._custom_templates, start=1):
            cb = QCheckBox(f"{i}、{filename}")
            template_path = _get_templates_dir() / filename
            cb.setToolTip(str(template_path))
            cb.toggled.connect(self._on_template_changed)
            check_layout.addWidget(cb)
            self.template_check_group.addButton(cb, id=i)
            self.custom_checks.append(cb)

        layout.addWidget(check_container)

        return group

    def _build_heading_mapping(self) -> QWidget:
        """标题映射展示（只读）。"""
        group = QGroupBox(
            "Markdown 标题 → Word 样式（固定映射，由 Pandoc 自动处理）"
        )
        layout = QHBoxLayout(group)
        layout.setSpacing(0)

        items = [
            ("#", "Heading 1"), ("##", "Heading 2"),
            ("###", "Heading 3"), ("####", "Heading 4"),
            ("#####", "Heading 5"), ("######", "Heading 6"),
        ]
        for md_pfx, word_style in items:
            item = self._make_label(f"{md_pfx}  →  {word_style}",
                                    object_name="descLabel", min_width=110)
            layout.addWidget(item)

        layout.addStretch()
        return group

    # ── 模板操作 ──────────────────────────────────────────

    def _open_builtin_template(self):
        """用系统默认程序打开内置模板。"""
        path = get_builtin_template_path()
        if path and path.is_file():
            try:
                os.startfile(str(path))
            except Exception:
                QMessageBox.warning(
                    self, "无法打开",
                    f"无法打开内置模板文件：\n{path}"
                )
        else:
            QMessageBox.information(
                self, "模板不存在",
                "内置模板文件（内置模板.docx）未找到。\n\n"
                "请确保 内置模板.docx 位于应用目录中。"
            )

    def _open_templates_folder(self):
        """用资源管理器打开模板文件夹。"""
        templates_dir = _get_templates_dir()
        try:
            os.startfile(str(templates_dir))
        except Exception:
            QMessageBox.warning(
                self, "无法打开",
                f"无法打开模板文件夹：\n{templates_dir}"
            )

    def _on_template_changed(self, _checked: bool):
        """模板单选按钮切换时触发。可用于实时预览状态更新。"""
        pass  # 当前仅用于保存时读取

    def _get_selected_template(self) -> str:
        """获取当前选中的模板配置值。"""
        checked_button = self.template_check_group.checkedButton()
        if checked_button is None or checked_button is self.builtin_check:
            return "built-in"

        # 自定义模板：从按钮文本解析文件名
        # 格式: "1、corporate-style.docx"
        text = checked_button.text()
        # 去掉编号前缀 "N、"
        for sep in ("、", ". "):
            if sep in text:
                filename = text.split(sep, 1)[-1].strip()
                break
        else:
            filename = text.strip()
        return f"custom:{filename}"

    def _set_selected_template(self, config_value: str):
        """根据配置值选中对应的模板复选框。"""
        mode, filename = _parse_template_config(config_value)

        if mode == "built-in" or mode == "missing":
            self.builtin_check.setChecked(True)
        elif mode == "custom" and filename:
            # 查找匹配的自定义按钮
            for cb in self.custom_checks:
                # 从按钮文本提取文件名
                text = cb.text()
                for sep in ("、", ". "):
                    if sep in text:
                        fname = text.split(sep, 1)[-1].strip()
                        break
                else:
                    fname = text.strip()
                if fname == filename:
                    cb.setChecked(True)
                    return
            # 文件已被删除，回退到内置
            self.builtin_check.setChecked(True)

    # ═══════════════════════════════════════════════════════
    #  Tab 2: Word → MD
    # ═══════════════════════════════════════════════════════

    def _build_word2md_tab(self) -> QWidget:
        """构建 Word→MD 标签页。"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(12)

        # ── 标题识别 ──
        layout.addWidget(self._build_heading_detection())

        # ── 内容提取 ──
        layout.addWidget(self._build_content_extraction())

        # ── 图片处理 ──
        layout.addWidget(self._build_image_settings())

        # ── 输出格式 ──
        layout.addWidget(self._build_output_format())

        layout.addStretch()

        return tab

    def _build_heading_detection(self) -> QWidget:
        """标题识别说明。"""
        group = QGroupBox("🔍 标题识别（基于大纲级别）")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        desc = self._make_label(
            "自动读取 Word 样式的大纲级别 (outlineLvl)，与样式名称无关：",
            object_name="descLabel", word_wrap=True)
        layout.addWidget(desc)

        # 映射展示行
        map_row = QHBoxLayout()
        map_row.setSpacing(4)
        levels = [
            ("Lv0", "#"), ("Lv1", "##"), ("Lv2", "###"),
            ("Lv3", "####"), ("Lv4", "#####"), ("Lv5", "######"),
        ]
        for lv, md in levels:
            item = self._make_label(f"{lv} → {md}",
                                    object_name="codeLabel", min_width=75)
            map_row.addWidget(item)
        map_row.addStretch()
        layout.addLayout(map_row)

        note = self._make_label(
            "不同样式名但相同大纲级别 → 输出同一 MD 标题层级。"
            "无需手动配置任何映射。",
            object_name="descLabel", word_wrap=True)
        layout.addWidget(note)

        return group

    def _build_content_extraction(self) -> QWidget:
        """内容提取选项。"""
        group = QGroupBox("内容提取")
        layout = QHBoxLayout(group)
        layout.setSpacing(24)

        self.preserve_tables_cb = QCheckBox("转换表格")
        self.preserve_lists_cb = QCheckBox("转换列表")
        self.preserve_links_cb = QCheckBox("保留超链接")

        layout.addWidget(self.preserve_tables_cb)
        layout.addWidget(self.preserve_lists_cb)
        layout.addWidget(self.preserve_links_cb)
        layout.addStretch()

        return group

    def _build_image_settings(self) -> QWidget:
        """图片处理设置。"""
        group = QGroupBox("图片处理")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)

        self.extract_images_cb = QCheckBox("提取 Word 文档中嵌入的图片")
        layout.addWidget(self.extract_images_cb)

        # 图片保存路径
        img_sub = QHBoxLayout()
        img_sub.setContentsMargins(24, 0, 0, 0)
        img_sub.setSpacing(8)

        img_sub.addWidget(self._make_label("保存到文件夹："))
        self.image_folder_input = self._protect_text_widget(QLineEdit())
        self.image_folder_input.setMinimumWidth(120)
        img_sub.addWidget(self.image_folder_input)
        img_sub.addStretch()

        layout.addLayout(img_sub)

        return group

    def _build_output_format(self) -> QWidget:
        """输出格式设置。"""
        group = QGroupBox("输出格式")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)

        self.add_toc_cb = QCheckBox("在 Markdown 开头添加 [TOC] 目录占位符")
        layout.addWidget(self.add_toc_cb)

        # 编码选择
        enc_row = QHBoxLayout()
        enc_row.setSpacing(8)

        enc_row.addWidget(self._make_label("输出编码："))
        self.encoding_combo = self._protect_text_widget(QComboBox())
        self.encoding_combo.addItems(["utf-8", "utf-8-sig", "gbk", "gb2312"])
        self.encoding_combo.setCurrentIndex(0)
        enc_row.addWidget(self.encoding_combo)
        enc_row.addStretch()

        layout.addLayout(enc_row)

        return group

    # ═══════════════════════════════════════════════════════
    #  Tab 3: 其他设置
    # ═══════════════════════════════════════════════════════

    def _build_other_tab(self) -> QWidget:
        """构建其他设置标签页。"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(12)

        # ── 通知设置 ──
        layout.addWidget(self._build_notification_settings())

        layout.addStretch()

        return tab

    def _build_notification_settings(self) -> QWidget:
        """Windows 系统通知设置。"""
        group = QGroupBox("🔔 Windows 系统通知")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)

        self.notif_enabled_cb = QCheckBox("启用通知（关闭后不显示任何通知）")
        layout.addWidget(self.notif_enabled_cb)

        self.show_success_cb = QCheckBox("转换成功时也通知（关闭后仅在出错时弹出）")
        layout.addWidget(self.show_success_cb)

        # 显示时长
        dur_row = QHBoxLayout()
        dur_row.setSpacing(8)
        dur_row.addWidget(self._make_label("显示时长："))
        self.duration_combo = self._protect_text_widget(QComboBox())
        self.duration_combo.addItems(["5秒", "10秒", "25秒", "持续"])
        self.duration_combo.setToolTip(
            "5秒/10秒/25秒 — 到时自动消失\n"
            "持续 — 一直显示直到手动关闭"
        )
        dur_row.addWidget(self.duration_combo)
        dur_row.addStretch()
        layout.addLayout(dur_row)

        self.show_path_cb = QCheckBox("显示输出文件路径（可点击打开所在文件夹）")
        layout.addWidget(self.show_path_cb)

        # 预览按钮
        preview_row = QHBoxLayout()
        preview_row.setSpacing(12)
        preview_btn = QPushButton("🔔 预览通知效果")
        preview_btn.clicked.connect(self._preview_notification)
        preview_row.addWidget(preview_btn)
        preview_label = self._make_label("查看当前设置下的通知效果",
                                         object_name="descLabel")
        preview_row.addWidget(preview_label)
        preview_row.addStretch()
        layout.addLayout(preview_row)

        return group

    # ── 通知预览 ──────────────────────────────────────────

    def _preview_notification(self):
        """显示通知预览。先隐藏主窗口再恢复。"""
        from notification import show_notification

        self.hide()
        # 延迟一下确保窗口已隐藏
        QApplication.processEvents()

        show_notification(
            title="✅ 转换成功（预览）",
            message="这是通知预览效果。\n\n"
                    "源文件：example.docx\n"
                    "输出文件：example.md",
            output_path=str(Path.home() / "Desktop" / "example.md"),
            auto_close=self.duration_combo.currentText(),
        )

        self.show()

    # ═══════════════════════════════════════════════════════
    #  配置读写
    # ═══════════════════════════════════════════════════════

    def _load_config_values(self):
        """将配置文件的值填入 UI 控件。"""
        # ── Tab 1: 模板选择 ──
        ref_doc = self.config.get("pandoc_reference_doc", "built-in")
        self._set_selected_template(ref_doc)

        # ── Tab 2: 内容提取 ──
        self.preserve_tables_cb.setChecked(
            self.config.get("preserve_tables", True))
        self.preserve_lists_cb.setChecked(
            self.config.get("preserve_lists", True))
        self.preserve_links_cb.setChecked(
            self.config.get("preserve_hyperlinks", True))

        # ── Tab 2: 图片处理 ──
        self.extract_images_cb.setChecked(
            self.config.get("extract_images", True))
        self.image_folder_input.setText(
            self.config.get("image_folder", "images"))

        # ── Tab 2: 输出格式 ──
        self.add_toc_cb.setChecked(
            self.config.get("add_toc", True))
        encoding = self.config.get("output_encoding", "utf-8")
        idx = self.encoding_combo.findText(encoding)
        if idx >= 0:
            self.encoding_combo.setCurrentIndex(idx)

        # ── Tab 3: 通知 ──
        notif = self.config.get("notification", {})
        self.notif_enabled_cb.setChecked(notif.get("enabled", True))
        self.show_success_cb.setChecked(notif.get("show_success", True))
        duration = notif.get("duration", "5秒")
        idx = self.duration_combo.findText(duration)
        if idx >= 0:
            self.duration_combo.setCurrentIndex(idx)
        else:
            # 兼容旧配置中的无效值（如已移除的"紧急"），回退到默认
            self.duration_combo.setCurrentIndex(0)  # "5秒"
            notif["duration"] = "5秒"  # 修正内存中的脏值，下次保存时自动清理
        self.show_path_cb.setChecked(notif.get("show_output_path", True))

    def _save_settings(self):
        """收集 UI 值并保存配置。"""
        # ── Tab 1: 模板 ──
        self.config["pandoc_reference_doc"] = self._get_selected_template()

        # ── Tab 2: 内容提取 ──
        self.config["preserve_tables"] = self.preserve_tables_cb.isChecked()
        self.config["preserve_lists"] = self.preserve_lists_cb.isChecked()
        self.config["preserve_hyperlinks"] = self.preserve_links_cb.isChecked()

        # ── Tab 2: 图片 ──
        self.config["extract_images"] = self.extract_images_cb.isChecked()
        self.config["image_folder"] = self.image_folder_input.text().strip() or "images"

        # ── Tab 2: 输出 ──
        self.config["add_toc"] = self.add_toc_cb.isChecked()
        self.config["output_encoding"] = self.encoding_combo.currentText()

        # ── Tab 3: 通知 ──
        notif = self.config.setdefault("notification", {})
        notif["enabled"] = self.notif_enabled_cb.isChecked()
        notif["show_success"] = self.show_success_cb.isChecked()
        notif["duration"] = self.duration_combo.currentText()
        notif["show_output_path"] = self.show_path_cb.isChecked()
        # 清理旧配置键
        notif.pop("auto_close_seconds", None)

        # 清理已不再使用的旧配置项
        self.config.pop("use_pandoc", None)
        heading = self.config.get("heading_mapping", {})
        heading.pop("word_to_md", None)

        if save_config(self.config):
            self._dirty = False
            QMessageBox.information(self, "保存成功", "设置已保存到配置文件。")
        else:
            QMessageBox.critical(
                self, "保存失败",
                "无法写入配置文件，请检查文件权限。"
            )

    def _reset_defaults(self):
        """恢复所有设置为默认值。"""
        reply = QMessageBox.question(
            self, "确认恢复默认",
            "确定要恢复所有设置为默认值吗？\n\n"
            "这将重置：\n"
            "  · 样式模板 → 内置模板\n"
            "  · 内容提取 → 全部开启\n"
            "  · 图片提取 → 开启\n"
            "  · 输出格式 → UTF-8，添加 TOC\n"
            "  · 通知 → 全部开启，5秒",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        from config import _DEFAULT_CONFIG
        _DEFAULT_CONFIG["notification"]["duration"] = "5秒"
        _DEFAULT_CONFIG["pandoc_reference_doc"] = "built-in"
        _DEFAULT_CONFIG["heading_mapping"] = {"md_to_word": {
            "#": "Heading 1", "##": "Heading 2", "###": "Heading 3",
            "####": "Heading 4", "#####": "Heading 5", "######": "Heading 6",
        }}
        save_config(_DEFAULT_CONFIG)

        # 从备份恢复内置模板（用户可能删除或修改了它）
        if restore_builtin_from_backup():
            # 模板已恢复 — 刷新内置复选框状态
            self._builtin_available = True
            self.builtin_template_path = get_builtin_template_path()
            self.builtin_check.setEnabled(True)
            self.builtin_check.setText("内置模板.docx（默认）")
            if self.builtin_template_path:
                self.builtin_check.setToolTip(str(self.builtin_template_path))

        # 重新加载配置并刷新 UI（无需销毁窗口）
        self.config = load_config()
        self._load_config_values()
        self._dirty = False


# ═══════════════════════════════════════════════════════════
#  入口
# ═══════════════════════════════════════════════════════════

def main():
    """启动设置窗口。"""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # 生成 checkmark 图标并注入 QSS
    icon_path = _ensure_checkmark_icon()

    # 确保内置模板备份存在（用于"恢复默认设置"还原）
    create_builtin_backup()
    stylesheet = STYLESHEET.replace(
        "CHECKMARK_PLACEHOLDER",
        str(icon_path).replace("\\", "/"))
    app.setStyleSheet(stylesheet)

    window = SettingsWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
