"""
PyQt6-based GUI with a modern Windows 11 tab layout.
Tab 1: Timeline Dashboard with embedded interactive chart
Tab 2: QCTP (Quality, Cost, Time, Performance)
Tab 3: Activities
Tab 4: Project Manager
Tab 5: Milestone & Phase Editor
Tab 6: Resources
Tab 7: Admin Panel (admin only)
"""

from __future__ import annotations

import datetime
import pathlib
import sys

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog, QVBoxLayout, QHBoxLayout,
    QGridLayout, QFormLayout, QTabWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QPushButton, QLabel, QLineEdit, QComboBox, QTextEdit,
    QGroupBox, QFrame, QScrollArea, QSplitter, QCheckBox, QRadioButton,
    QButtonGroup, QMessageBox, QFileDialog, QInputDialog, QMenu,
    QMenuBar, QStatusBar, QToolBar, QSizePolicy, QSpacerItem, QListWidget,
    QListWidgetItem, QAbstractItemView, QTreeWidget, QTreeWidgetItem,
    QDialogButtonBox, QProgressBar, QSlider, QSpinBox, QDoubleSpinBox,
    QDateEdit, QTimeEdit, QDateTimeEdit, QCalendarWidget, QToolButton
)
from PyQt6.QtCore import (
    Qt, QTimer, QSize, pyqtSignal, QDate, QTime, QDateTime, 
    QPoint, QRect, QMargins
)
from PyQt6.QtGui import (
    QFont, QColor, QPalette, QIcon, QPixmap, QAction, QKeySequence,
    QFontMetrics, QPainter, QBrush, QPen, QCursor
)

import matplotlib
matplotlib.use("QtAgg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

from timeline_tool import database as db
from timeline_tool import auth
from timeline_tool import config as cfg
from timeline_tool.models import Project, ReferenceLine
from timeline_tool.utils import date_range_padded
from timeline_tool.renderer import render_timeline

from timeline_tool.export_report import (
    generate_pdf_report, generate_excel_report,
    check_dependencies, REPORTLAB_AVAILABLE, OPENPYXL_AVAILABLE
)

from timeline_tool.resources import (
    Resource, init_resource_tables, add_resource, update_resource,
    delete_resource, get_all_resources, assign_resource_to_project,
    remove_assignment, get_project_assignments, get_resource_assignments,
    calculate_resource_utilization, get_team_utilization_summary,
    find_available_resources
)

from timeline_tool.audit_viewer import (
    get_audit_log, get_unique_users, get_unique_actions,
    get_activity_summary, export_audit_log_csv, get_entity_history,
    get_action_description, get_action_icon, ACTION_DESCRIPTIONS
)

from timeline_tool.backup import (
    create_backup, restore_backup, list_backups, delete_backup,
    export_to_json, import_from_json, get_backup_dir
)

# ─────────────────────────────────────────────────────────────────────────
# Phase colours
# ─────────────────────────────────────────────────────────────────────────
PHASE_PALETTE = ["#85C1E9", "#82E0AA", "#F8C471", "#D7BDE2", "#F0B27A", "#AED6F1"]

REGION_OPTIONS = ["", "IAP", "EMEA", "NAFTA", "LATAM", "ROW"]


def _pick_color(index: int, project: Project) -> str:
    if project.color:
        return project.color
    return cfg.DEFAULT_COLORS[index % len(cfg.DEFAULT_COLORS)]


def _status_edge_color(project: Project) -> str:
    return cfg.STATUS_COLORS.get(project.computed_status(), cfg.STATUS_COLORS["on-track"])


# ─────────────────────────────────────────────────────────────────────────
# Windows 11 style QSS
# ─────────────────────────────────────────────────────────────────────────

WIN11_STYLE = """
QWidget {
    background-color: #F3F3F3;
    color: #1A1A1A;
    font-family: 'Segoe UI';
    font-size: 10pt;
}
QMainWindow {
    background-color: #F3F3F3;
}
QPushButton {
    background-color: #FFFFFF;
    color: #1A1A1A;
    border: 1px solid #E0E0E0;
    border-radius: 4px;
    padding: 6px 12px;
    font-size: 10pt;
}
QPushButton:hover {
    background-color: #E8F0FE;
}
QPushButton:pressed {
    background-color: #0067C0;
    color: white;
}
QPushButton#accentBtn {
    background-color: #0067C0;
    color: white;
    font-weight: bold;
    padding: 8px 14px;
}
QPushButton#accentBtn:hover {
    background-color: #005A9E;
}
QTabWidget::pane {
    border: 1px solid #E0E0E0;
    background-color: #FFFFFF;
}
QTabBar::tab {
    background-color: #E8E8E8;
    color: #1A1A1A;
    padding: 10px 20px;
    font-size: 11pt;
}
QTabBar::tab:selected {
    background-color: #FFFFFF;
    color: #0067C0;
}
QTableWidget {
    background-color: #FFFFFF;
    alternate-background-color: #F8F8F8;
    gridline-color: #E0E0E0;
    font-size: 10pt;
}
QTableWidget::item:selected {
    background-color: #E8F0FE;
    color: #1A1A1A;
}
QHeaderView::section {
    background-color: #F3F3F3;
    color: #1A1A1A;
    font-weight: bold;
    padding: 4px;
    border: 1px solid #E0E0E0;
}
QLineEdit {
    background-color: #F5F5F5;
    border: 1px solid #DDDDDD;
    border-radius: 4px;
    padding: 6px;
    font-size: 10pt;
}
QComboBox {
    background-color: #FFFFFF;
    border: 1px solid #DDDDDD;
    border-radius: 4px;
    padding: 4px 8px;
    font-size: 10pt;
}
QGroupBox {
    font-weight: bold;
    border: 1px solid #E0E0E0;
    border-radius: 4px;
    margin-top: 8px;
    padding-top: 8px;
    background-color: #FFFFFF;
}
QGroupBox::title {
    color: #0067C0;
    subcontrol-origin: margin;
    left: 10px;
}
QTextEdit {
    background-color: #FFFFFF;
    border: 1px solid #DDDDDD;
    border-radius: 4px;
    padding: 4px;
    font-size: 10pt;
}
QScrollArea {
    border: none;
    background-color: transparent;
}
QTreeWidget {
    background-color: #FFFFFF;
    alternate-background-color: #F8F8F8;
    border: 1px solid #E0E0E0;
}
QTreeWidget::item:selected {
    background-color: #E8F0FE;
    color: #1A1A1A;
}
QListWidget {
    background-color: #FFFFFF;
    border: 1px solid #DDDDDD;
    border-radius: 4px;
}
QListWidget::item:selected {
    background-color: #E8F0FE;
    color: #1A1A1A;
}
QDateEdit {
    background-color: #FFFFFF;
    border: 1px solid #DDDDDD;
    border-radius: 4px;
    padding: 4px;
}
QSpinBox, QDoubleSpinBox {
    background-color: #FFFFFF;
    border: 1px solid #DDDDDD;
    border-radius: 4px;
    padding: 4px;
}
QDialog {
    background-color: #F3F3F3;
}
"""


# ─────────────────────────────────────────────────────────────────────────
# Login Window
# ─────────────────────────────────────────────────────────────────────────

class LoginWindow(QDialog):
    def __init__(self, db_path: pathlib.Path | None = None, parent=None):
        super().__init__(parent)
        self.db_path = db_path
        self.user = None
        self._password_visible = False

        self.setWindowTitle("Project Dashboard")
        self.setFixedSize(450, 650)
        self.setModal(True)

        self._setup_ui()
        self._load_remembered_user()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(40, 25, 40, 25)
        main_layout.setSpacing(0)

        # Logo placeholder
        logo_label = QLabel("Project")
        logo_label.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(logo_label)

        # Title & Subtitle
        title_label = QLabel("Project Dashboard")
        title_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)

        subtitle_label = QLabel("Sign in to access your projects")
        subtitle_label.setFont(QFont("Segoe UI", 10))
        subtitle_label.setStyleSheet("color: #666666;")
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(subtitle_label)
        main_layout.addSpacing(18)

        # Login Card
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #E0E0E0;
                border-radius: 4px;
            }
        """)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(25, 20, 25, 20)
        card_layout.setSpacing(12)

        # Username Field
        username_label = QLabel("Username")
        username_label.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        card_layout.addWidget(username_label)

        self.username_entry = QLineEdit()
        self.username_entry.setPlaceholderText("Enter your username")
        self.username_entry.setMinimumHeight(40)
        self.username_entry.returnPressed.connect(lambda: self.password_entry.setFocus())
        card_layout.addWidget(self.username_entry)

        # Password Field
        password_label = QLabel("Password")
        password_label.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        card_layout.addWidget(password_label)

        password_container = QHBoxLayout()
        self.password_entry = QLineEdit()
        self.password_entry.setPlaceholderText("Enter your password")
        self.password_entry.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_entry.setMinimumHeight(40)
        self.password_entry.returnPressed.connect(self._login)
        password_container.addWidget(self.password_entry)

        self.eye_btn = QPushButton("👁")
        self.eye_btn.setFixedSize(40, 40)
        self.eye_btn.setStyleSheet("border: none; background: transparent;")
        self.eye_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.eye_btn.clicked.connect(self._toggle_password_visibility)
        password_container.addWidget(self.eye_btn)
        card_layout.addLayout(password_container)

        # Remember Me & Forgot Password
        remember_layout = QHBoxLayout()
        self.remember_check = QCheckBox("Remember me")
        self.remember_check.setStyleSheet("color: #666666;")
        remember_layout.addWidget(self.remember_check)

        forgot_label = QLabel('<a href="#" style="color: #0067C0;">Forgot password?</a>')
        forgot_label.setCursor(Qt.CursorShape.PointingHandCursor)
        forgot_label.mousePressEvent = self._on_forgot_password
        remember_layout.addWidget(forgot_label)
        card_layout.addLayout(remember_layout)

        # Sign In Button
        self.signin_btn = QPushButton("Sign In")
        self.signin_btn.setObjectName("accentBtn")
        self.signin_btn.setMinimumHeight(48)
        self.signin_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.signin_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.signin_btn.clicked.connect(self._login)
        card_layout.addWidget(self.signin_btn)

        main_layout.addWidget(card)

        # Status Label
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: red;")
        self.status_label.setFont(QFont("Segoe UI", 9))
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setWordWrap(True)
        main_layout.addWidget(self.status_label)
        main_layout.addSpacing(10)

        # Help Text
        help_label = QLabel("Need help? Contact your administrator")
        help_label.setStyleSheet("color: #999999;")
        help_label.setFont(QFont("Segoe UI", 9))
        help_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(help_label)

        main_layout.addStretch()

        # Footer
        footer_layout = QVBoxLayout()
        footer_sep = QFrame()
        footer_sep.setFrameShape(QFrame.Shape.HLine)
        footer_sep.setStyleSheet("background-color: #E0E0E0;")
        footer_layout.addWidget(footer_sep)

        dev_label = QLabel("Developed by IAP VSSQ AI/ML Team")
        dev_label.setFont(QFont("Segoe UI", 9, italic=True))
        dev_label.setStyleSheet("color: #888888;")
        dev_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_layout.addWidget(dev_label)

        version_label = QLabel("v2.0 | © 2025")
        version_label.setFont(QFont("Segoe UI", 8))
        version_label.setStyleSheet("color: #AAAAAA;")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_layout.addWidget(version_label)

        main_layout.addLayout(footer_layout)

        self.username_entry.setFocus()

    def _toggle_password_visibility(self):
        self._password_visible = not self._password_visible
        if self._password_visible:
            self.password_entry.setEchoMode(QLineEdit.EchoMode.Normal)
            self.eye_btn.setText("🙈")
        else:
            self.password_entry.setEchoMode(QLineEdit.EchoMode.Password)
            self.eye_btn.setText("👁")

    def _on_forgot_password(self, event=None):
        QMessageBox.information(
            self,
            "Reset Password",
            "To reset your password, please contact your system administrator.\n\n"
            "Email: admin@company.com\n"
            "Phone: ext. 1234\n\n"
            "Please have your username ready when you contact support."
        )

    def _load_remembered_user(self):
        try:
            remember_file = pathlib.Path.home() / ".project_dashboard_remember"
            if remember_file.exists():
                username = remember_file.read_text().strip()
                if username:
                    self.username_entry.setText(username)
                    self.remember_check.setChecked(True)
                    self.password_entry.setFocus()
        except Exception:
            pass

    def _save_remembered_user(self, username: str):
        try:
            remember_file = pathlib.Path.home() / ".project_dashboard_remember"
            if self.remember_check.isChecked() and username:
                remember_file.write_text(username)
            elif remember_file.exists():
                remember_file.unlink()
        except Exception:
            pass

    def _login(self):
        username = self.username_entry.text().strip()
        password = self.password_entry.text().strip()

        if not username:
            self.status_label.setText("Please enter your username.")
            self.username_entry.setFocus()
            return

        if not password:
            self.status_label.setText("Please enter your password.")
            self.password_entry.setFocus()
            return

        # Show loading state
        original_text = self.signin_btn.text()
        self.signin_btn.setText("Signing in...")
        self.signin_btn.setEnabled(False)
        QApplication.processEvents()

        user = auth.authenticate(username, password, self.db_path)

        if user is None:
            self.signin_btn.setText(original_text)
            self.signin_btn.setEnabled(True)
            self.status_label.setText("Invalid username or password. Please try again.")
            self.password_entry.clear()
            self.password_entry.setFocus()

            # Flash the button red briefly
            self.signin_btn.setStyleSheet("background-color: #DC2626; color: white;")
            QTimer.singleShot(200, lambda: self.signin_btn.setStyleSheet(""))
            return

        # Save remember-me preference
        self._save_remembered_user(username)

        self.user = user
        self.accept()


# ─────────────────────────────────────────────────────────────────────────
# Main Application (Tab Layout)
# ─────────────────────────────────────────────────────────────────────────

class MainApp(QMainWindow):
    def __init__(self, user: dict, db_path: pathlib.Path | None = None):
        super().__init__()
        self.user = user
        self.db_path = db_path
        self.can_edit = user["permissions"]["can_edit"]
        self.can_manage = user["permissions"]["can_manage_users"]

        display_name = user.get("full_name") or user["username"]
        role_display = user["role"].upper()
        self.setWindowTitle("Project Dashboard")
        self.setMinimumSize(1000, 650)

        self._setup_ui(display_name, role_display)
        self.showMaximized()

    def _setup_ui(self, display_name: str, role_display: str):
        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(0)

        # Top bar
        topbar = QFrame()
        topbar.setStyleSheet("background-color: #F3F3F3; padding: 8px 15px;")
        topbar_layout = QHBoxLayout(topbar)
        topbar_layout.setContentsMargins(15, 8, 15, 8)

        title_label = QLabel("📊 Project Dashboard")
        title_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        topbar_layout.addWidget(title_label)

        topbar_layout.addStretch()

        # User pill
        user_label = QLabel(f"👤 {display_name}")
        user_label.setFont(QFont("Segoe UI", 10))
        topbar_layout.addWidget(user_label)

        role_colors = {"admin": "#0067C0", "editor": "#16A34A", "viewer": "#9333EA"}
        role_color = role_colors.get(self.user["role"], "#666666")
        role_badge = QLabel(f" {role_display} ")
        role_badge.setStyleSheet(f"""
            background-color: {role_color};
            color: white;
            font-weight: bold;
            font-size: 8pt;
            padding: 2px 8px;
            border-radius: 3px;
        """)
        topbar_layout.addWidget(role_badge)

        main_layout.addWidget(topbar)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("background-color: #E0E0E0;")
        main_layout.addWidget(sep)

        # Notebook (tabs)
        self.notebook = QTabWidget()
        self.notebook.setDocumentMode(True)

        # Tab 1: Dashboard
        self.tab_dashboard = QWidget()
        self.notebook.addTab(self.tab_dashboard, "  📊 Dashboard  ")
        self._build_dashboard_tab()

        # Tab 2: QCTP
        self.tab_qctp = QWidget()
        self.notebook.addTab(self.tab_qctp, "  📋 QCTP  ")
        self._build_qctp_tab()

        # Tab 3: Activities
        self.tab_activities = QWidget()
        self.notebook.addTab(self.tab_activities, "  📝 Activities  ")
        self._build_activities_tab()

        # Tab 4: Projects
        self.tab_projects = QWidget()
        self.notebook.addTab(self.tab_projects, "  📁 Projects  ")
        self._build_projects_tab()

        # Tab 5: Milestones & Phases
        self.tab_milestones = QWidget()
        self.notebook.addTab(self.tab_milestones, "  🎯 Milestones & Phases  ")
        self._build_milestones_tab()

        # Tab 6: Resources
        self.tab_resources = QWidget()
        self.notebook.addTab(self.tab_resources, "  👥 Resources  ")
        self._build_resources_tab()

        # Tab 7: Admin (admin only)
        if self.can_manage:
            self.tab_admin = QWidget()
            self.notebook.addTab(self.tab_admin, "  ⚙️ Admin  ")
            self._build_admin_tab()

        main_layout.addWidget(self.notebook)

        # Status bar
        status_bar = QFrame()
        status_bar.setStyleSheet("background-color: #F3F3F3; padding: 3px 10px;")
        status_layout = QHBoxLayout(status_bar)
        status_layout.setContentsMargins(10, 3, 10, 3)

        self.status_text = QLabel("Ready")
        self.status_text.setFont(QFont("Segoe UI", 9))
        self.status_text.setStyleSheet("color: #666666;")
        status_layout.addWidget(self.status_text)

        status_layout.addStretch()

        self.coords_label = QLabel("")
        self.coords_label.setFont(QFont("Segoe UI", 8))
        self.coords_label.setStyleSheet("color: #999999;")
        status_layout.addWidget(self.coords_label)

        main_layout.addWidget(status_bar)

        # Footer
        footer_sep = QFrame()
        footer_sep.setFrameShape(QFrame.Shape.HLine)
        footer_sep.setStyleSheet("background-color: #E0E0E0;")
        main_layout.addWidget(footer_sep)

        footer = QFrame()
        footer.setStyleSheet("background-color: #F3F3F3; padding: 6px 10px;")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(10, 6, 10, 6)

        dev_label = QLabel("Developed by IAP VSSQ AI/ML Team")
        dev_label.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        dev_label.setStyleSheet("color: #555555;")
        footer_layout.addWidget(dev_label)

        footer_layout.addStretch()

        db_label = QLabel(f"Database: {self.db_path or db.DEFAULT_DB_PATH}")
        db_label.setFont(QFont("Segoe UI", 8))
        db_label.setStyleSheet("color: #999999;")
        footer_layout.addWidget(db_label)

        main_layout.addWidget(footer)

    def _set_status(self, text: str):
        self.status_text.setText(text)
        QTimer.singleShot(5000, lambda: self.status_text.setText("Ready"))

    # ═════════════════════════════════════════════════════════════════════
    # TAB 1 — DASHBOARD (Embedded Chart)
    # ═════════════════════════════════════════════════════════════════════

    def _build_dashboard_tab(self):
        layout = QVBoxLayout(self.tab_dashboard)
        layout.setContentsMargins(3, 3, 3, 3)
        layout.setSpacing(5)

        # Top row: Summary cards + Action buttons
        top_row = QHBoxLayout()

        # Summary cards
        self.cards_frame = QWidget()
        self.cards_layout = QHBoxLayout(self.cards_frame)
        self.cards_layout.setContentsMargins(0, 0, 0, 0)
        self.cards_layout.setSpacing(5)
        top_row.addWidget(self.cards_frame, stretch=1)

        # Buttons
        btn_frame = QHBoxLayout()
        btn_frame.setSpacing(3)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_dashboard)
        btn_frame.addWidget(refresh_btn)

        popout_btn = QPushButton("🪟 Pop-out Chart")
        popout_btn.clicked.connect(self._popout_timeline)
        btn_frame.addWidget(popout_btn)

        save_png_btn = QPushButton("💾 Save PNG")
        save_png_btn.clicked.connect(self._save_timeline_png)
        btn_frame.addWidget(save_png_btn)

        save_html_btn = QPushButton("🌐 Save HTML")
        save_html_btn.clicked.connect(self._save_timeline_html)
        btn_frame.addWidget(save_html_btn)

        pdf_btn = QPushButton("📄 Export PDF")
        pdf_btn.clicked.connect(self._export_pdf_report)
        btn_frame.addWidget(pdf_btn)

        excel_btn = QPushButton("📊 Export Excel")
        excel_btn.clicked.connect(self._export_excel_report)
        btn_frame.addWidget(excel_btn)

        summary_btn = QPushButton("📈 Summary")
        summary_btn.clicked.connect(self._show_summary_dashboard)
        btn_frame.addWidget(summary_btn)

        # Theme toggle
        self.theme_var = "light"
        light_radio = QRadioButton("Light")
        light_radio.setChecked(True)
        light_radio.toggled.connect(lambda checked: self._set_theme("light") if checked else None)
        btn_frame.addWidget(light_radio)

        dark_radio = QRadioButton("Dark")
        dark_radio.toggled.connect(lambda checked: self._set_theme("dark") if checked else None)
        btn_frame.addWidget(dark_radio)

        top_row.addLayout(btn_frame)
        layout.addLayout(top_row)

        # Filter bar
        filter_group = QGroupBox("  🔍 Filters & Controls  ")
        filter_layout = QHBoxLayout(filter_group)
        filter_layout.setSpacing(10)

        # Status filter
        filter_layout.addWidget(QLabel("Status:"))
        self.filter_status = QComboBox()
        self.filter_status.addItems(["All", "On Track", "At Risk", "Overdue"])
        self.filter_status.currentTextChanged.connect(self._refresh_dashboard)
        filter_layout.addWidget(self.filter_status)

        # Dev Region filter
        filter_layout.addWidget(QLabel("Dev Region:"))
        self.filter_dev_region = QComboBox()
        self.filter_dev_region.addItems(["All", "IAP", "EMEA", "NAFTA", "LATAM", "ROW"])
        self.filter_dev_region.currentTextChanged.connect(self._refresh_dashboard)
        filter_layout.addWidget(self.filter_dev_region)

        # Sales Region filter
        filter_layout.addWidget(QLabel("Sales Region:"))
        self.filter_sales_region = QComboBox()
        self.filter_sales_region.addItems(["All", "IAP", "EMEA", "NAFTA", "LATAM", "ROW"])
        self.filter_sales_region.currentTextChanged.connect(self._refresh_dashboard)
        filter_layout.addWidget(self.filter_sales_region)

        # Quarter filter
        filter_layout.addWidget(QLabel("Quarter:"))
        self.filter_quarter = QComboBox()
        self.filter_quarter.addItems(["All", "Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025",
                                       "Q1 2026", "Q2 2026", "Q3 2026", "Q4 2026"])
        self.filter_quarter.currentTextChanged.connect(self._refresh_dashboard)
        filter_layout.addWidget(self.filter_quarter)

        # Search
        filter_layout.addWidget(QLabel("Search:"))
        self.filter_search = QLineEdit()
        self.filter_search.setPlaceholderText("Project name...")
        self.filter_search.setMaximumWidth(200)
        self.filter_search.returnPressed.connect(self._refresh_dashboard)
        self.filter_search.textChanged.connect(self._on_search_keyrelease)
        filter_layout.addWidget(self.filter_search)

        apply_btn = QPushButton("🔍 Apply")
        apply_btn.clicked.connect(self._refresh_dashboard)
        filter_layout.addWidget(apply_btn)

        clear_btn = QPushButton("🔄 Clear")
        clear_btn.clicked.connect(self._clear_filters)
        filter_layout.addWidget(clear_btn)

        zoom_btn = QPushButton("📐 Zoom to Fit")
        zoom_btn.clicked.connect(self._zoom_to_fit)
        filter_layout.addWidget(zoom_btn)

        filter_layout.addStretch()
        layout.addWidget(filter_group)

        # Chart area
        self.chart_frame = QWidget()
        self.chart_layout = QVBoxLayout(self.chart_frame)
        self.chart_layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.chart_frame, stretch=1)

        self.chart_canvas = None
        self.chart_toolbar = None
        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._refresh_dashboard)

        self._refresh_dashboard()

    def _set_theme(self, theme: str):
        self.theme_var = theme
        self._refresh_dashboard()

    def _on_search_keyrelease(self):
        self._search_timer.start(500)

    def _zoom_to_fit(self):
        if self.chart_canvas:
            self._refresh_dashboard()
            self._set_status("Zoomed to fit all projects")

    def _clear_filters(self):
        self.filter_status.setCurrentText("All")
        self.filter_dev_region.setCurrentText("All")
        self.filter_sales_region.setCurrentText("All")
        self.filter_quarter.setCurrentText("All")
        self.filter_search.clear()
        self._refresh_dashboard()

    def _get_quarter_date_range(self, quarter_str: str) -> tuple:
        if quarter_str == "All" or not quarter_str:
            return None, None

        parts = quarter_str.split()
        if len(parts) != 2:
            return None, None

        q = parts[0]
        year = int(parts[1])

        quarter_starts = {"Q1": (1, 1), "Q2": (4, 1), "Q3": (7, 1), "Q4": (10, 1)}
        quarter_ends = {"Q1": (3, 31), "Q2": (6, 30), "Q3": (9, 30), "Q4": (12, 31)}

        if q not in quarter_starts:
            return None, None

        start_month, start_day = quarter_starts[q]
        end_month, end_day = quarter_ends[q]

        start_date = datetime.date(year, start_month, start_day)
        end_date = datetime.date(year, end_month, end_day)

        return start_date, end_date

    def _refresh_dashboard(self):
        projects, ref_lines = db.load_all(self.db_path)
        today = datetime.date.today()

        # Apply filters
        status_filter = self.filter_status.currentText()
        if status_filter and status_filter != "All":
            status_map = {"On Track": "on-track", "At Risk": "at-risk", "Overdue": "overdue"}
            target_status = status_map.get(status_filter)
            if target_status:
                projects = [p for p in projects if p.computed_status(today) == target_status]

        dev_filter = self.filter_dev_region.currentText()
        if dev_filter and dev_filter != "All":
            projects = [p for p in projects if p.dev_region == dev_filter]

        sales_filter = self.filter_sales_region.currentText()
        if sales_filter and sales_filter != "All":
            projects = [p for p in projects if p.sales_region == sales_filter]

        quarter_filter = self.filter_quarter.currentText()
        if quarter_filter and quarter_filter != "All":
            q_start, q_end = self._get_quarter_date_range(quarter_filter)
            if q_start and q_end:
                projects = [p for p in projects
                           if not (p.end_date < q_start or p.start_date > q_end)]

        search_term = self.filter_search.text().strip().lower()
        if search_term:
            projects = [p for p in projects if search_term in p.name.lower()]

        # Update summary cards
        self._update_summary_cards(projects, today)

        # Render chart
        if not projects:
            self._clear_chart()
            no_data_label = QLabel("No projects match the current filters.\nAdjust filters or add projects in the Projects tab.")
            no_data_label.setFont(QFont("Segoe UI", 14))
            no_data_label.setStyleSheet("color: #999999;")
            no_data_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.chart_layout.addWidget(no_data_label)
            self._set_status("Dashboard refreshed — no matching projects")
            return

        self._render_embedded_chart(projects, ref_lines)

        filter_info = ""
        if status_filter != "All":
            filter_info += f" | Status: {status_filter}"
        if dev_filter != "All":
            filter_info += f" | Dev: {dev_filter}"
        if sales_filter != "All":
            filter_info += f" | Sales: {sales_filter}"
        if quarter_filter != "All":
            filter_info += f" | {quarter_filter}"
        if search_term:
            filter_info += f" | Search: '{search_term}'"

        self._set_status(f"Dashboard refreshed — {len(projects)} project(s) loaded{filter_info}")

    def _clear_chart(self):
        while self.chart_layout.count():
            item = self.chart_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.chart_canvas = None
        self.chart_toolbar = None

    def _update_summary_cards(self, projects, today):
        # Clear existing cards
        while self.cards_layout.count():
            item = self.cards_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        total = len(projects)
        on_track = sum(1 for p in projects if p.computed_status(today) == "on-track")
        at_risk = sum(1 for p in projects if p.computed_status(today) == "at-risk")
        overdue = sum(1 for p in projects if p.computed_status(today) == "overdue")

        cards_data = [
            ("📁", "Total Projects", str(total), "#0067C0"),
            ("🟢", "On Track", str(on_track), "#16A34A"),
            ("🟡", "At Risk", str(at_risk), "#EAB308"),
            ("🔴", "Overdue", str(overdue), "#DC2626"),
        ]

        for icon, label, value, color in cards_data:
            card = QFrame()
            card.setStyleSheet(f"""
                QFrame {{
                    background-color: white;
                    border: 1px solid #E0E0E0;
                    border-radius: 4px;
                    padding: 6px 10px;
                }}
            """)
            card_layout = QHBoxLayout(card)
            card_layout.setContentsMargins(8, 6, 8, 6)

            icon_label = QLabel(icon)
            icon_label.setFont(QFont("Segoe UI", 16))
            card_layout.addWidget(icon_label)

            text_widget = QWidget()
            text_layout = QVBoxLayout(text_widget)
            text_layout.setContentsMargins(0, 0, 0, 0)
            text_layout.setSpacing(0)

            value_label = QLabel(value)
            value_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
            value_label.setStyleSheet(f"color: {color};")
            text_layout.addWidget(value_label)

            name_label = QLabel(label)
            name_label.setFont(QFont("Segoe UI", 8))
            name_label.setStyleSheet("color: #666666;")
            text_layout.addWidget(name_label)

            card_layout.addWidget(text_widget)
            self.cards_layout.addWidget(card)

    def _render_embedded_chart(self, projects: list, ref_lines: list):
        """Render the full Gantt chart embedded inside the Dashboard tab."""
        self._clear_chart()

        today = datetime.date.today()
        theme_cfg = cfg.THEMES.get(self.theme_var, cfg.THEMES["light"])
        n = len(projects)

        # Create figure
        MAX_VISIBLE = 6
        fig_h = max(4.0, min(n, MAX_VISIBLE) * 0.9 + 1.5)
        fig, ax = plt.subplots(figsize=(16, fig_h))
        fig.patch.set_facecolor(theme_cfg["bg_color"])
        ax.set_facecolor(theme_cfg["axis_bg"])
        ax.title.set_color(theme_cfg["text_color"])
        ax.xaxis.label.set_color(theme_cfg["text_color"])
        ax.tick_params(colors=theme_cfg["text_color"])
        for spine in ax.spines.values():
            spine.set_color(theme_cfg["grid_color"])

        x_min, x_max = date_range_padded(projects)
        y_min_orig = 0
        y_max_orig = n + 1

        y_positions = []
        y_labels = []
        milestone_artists = []
        project_bars = []

        for idx, project in enumerate(projects):
            y = n - idx
            y_positions.append(y)
            y_labels.append(project.name)
            color = _pick_color(idx, project)
            status_color = cfg.STATUS_COLORS.get(project.computed_status(today), cfg.STATUS_COLORS["on-track"])

            # Full bar (background)
            bar = ax.barh(y, (project.end_date - project.start_date).days,
                    left=project.start_date, height=cfg.BAR_HEIGHT,
                    color=color, alpha=0.25, edgecolor="none", linewidth=0, picker=True)

            project_bars.append({
                "bar": bar[0], "project": project, "y": y,
                "y_min": y - cfg.BAR_HEIGHT / 2, "y_max": y + cfg.BAR_HEIGHT / 2,
            })

            # Progress overlay
            progress = project.progress(today)
            if progress > 0:
                elapsed_days = (project.end_date - project.start_date).days * progress
                ax.barh(y, elapsed_days, left=project.start_date,
                        height=cfg.BAR_HEIGHT, color=color, alpha=0.85,
                        edgecolor="none", linewidth=0, zorder=3)

            # Phase sub-bars
            for p_idx, phase in enumerate(project.phases):
                phase_color = PHASE_PALETTE[p_idx % len(PHASE_PALETTE)]
                phase_days = (phase.end_date - phase.start_date).days
                ax.barh(y, phase_days, left=phase.start_date,
                        height=cfg.PHASE_HEIGHT, color=phase_color, alpha=0.75,
                        edgecolor="#666666", linewidth=0.5, zorder=3)
                if phase_days > 45:
                    mid = phase.start_date + datetime.timedelta(days=phase_days / 2)
                    ax.text(mid, y - cfg.BAR_HEIGHT / 2 - 0.05, phase.name,
                            fontsize=7, ha="center", va="top",
                            color="#333333", fontweight="bold", zorder=4)

            # Date labels
            ax.text(project.start_date, y - cfg.BAR_HEIGHT / 2 - 0.18,
                    project.start_date.strftime("%b %d, %Y"),
                    fontsize=7, ha="left", va="top", color=theme_cfg["date_label_color"])
            ax.text(project.end_date, y - cfg.BAR_HEIGHT / 2 - 0.18,
                    project.end_date.strftime("%b %d, %Y"),
                    fontsize=7, ha="right", va="top", color=theme_cfg["date_label_color"])

            # Milestones
            for ms_idx, ms in enumerate(project.milestones):
                milestone_color = ms.marker_color(today)
                sc = ax.scatter(ms.date, y, marker=cfg.MILESTONE_MARKER,
                                s=cfg.MILESTONE_SIZE, color=milestone_color,
                                edgecolors=cfg.MILESTONE_EDGE_COLOR,
                                linewidths=0.8, zorder=5, picker=5)
                milestone_artists.append((sc, project.name, ms))

                if ms_idx % 2 == 0:
                    label_y = y + cfg.BAR_HEIGHT * 0.6
                    va = "bottom"
                else:
                    label_y = y - cfg.BAR_HEIGHT * 0.6
                    va = "top"

                ax.text(mdates.date2num(ms.date), label_y, ms.name,
                        fontsize=6, ha="center", va=va,
                        color=theme_cfg["text_color"], fontweight="bold", zorder=7)

        # Today line
        ax.axvline(today, color=cfg.TODAY_LINE_COLOR,
                linewidth=cfg.TODAY_LINE_WIDTH, linestyle=cfg.TODAY_LINE_STYLE,
                label=f"Today ({today.strftime('%b %d, %Y')})", zorder=4)

        ax.text(mdates.date2num(today), y_max_orig - 0.05,
            today.strftime("%b %d, %Y"),
            fontsize=7, fontweight="bold", color=cfg.TODAY_LINE_COLOR,
            ha="center", va="top", zorder=8,
            bbox=dict(boxstyle="round,pad=0.2", facecolor="white",
                      edgecolor=cfg.TODAY_LINE_COLOR, alpha=0.85))

        # Reference lines
        for ref in ref_lines:
            ax.axvline(ref.date, color=ref.color, linewidth=1.8,
                       linestyle=ref.style, zorder=4)
            ax.text(ref.date, y_max_orig - 0.15,
                    f"  {ref.name}\n  {ref.date.strftime('%b %d, %Y')}",
                    fontsize=7, color=ref.color, fontweight="bold",
                    ha="left", va="top")

        # Axes setup
        ax.set_xlim(x_min, x_max)
        if n > MAX_VISIBLE:
            y_view_min = n - MAX_VISIBLE + 0.5
            y_view_max = n + 1
        else:
            y_view_min = y_min_orig
            y_view_max = y_max_orig
        ax.set_ylim(y_view_min, y_view_max)
        ax.set_yticks(y_positions)
        ax.set_yticklabels([""] * len(y_positions))

        # Custom y-axis labels
        for idx, project in enumerate(projects):
            y = y_positions[idx]
            status_text = project.computed_status(today).replace("-", " ").upper()
            s_color = cfg.STATUS_COLORS.get(project.computed_status(today), cfg.STATUS_COLORS["on-track"])

            ax.text(-0.01, y + 0.12, project.name,
                    transform=ax.get_yaxis_transform(),
                    fontsize=9, ha="right", va="center",
                    color=theme_cfg["text_color"], fontweight="bold", clip_on=False)
            ax.text(-0.01, y - 0.18, f"● {status_text}",
                    transform=ax.get_yaxis_transform(),
                    fontsize=7, ha="right", va="center",
                    color=s_color, fontweight="bold", clip_on=False)

        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.xaxis.set_minor_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
        fig.autofmt_xdate(rotation=45, ha="right")
        ax.grid(axis="x", linestyle=":", linewidth=0.5, alpha=0.6, color=theme_cfg["grid_color"])
        ax.set_axisbelow(True)
        ax.set_title("Project Timelines", fontsize=14, fontweight="bold",
                      pad=15, color=theme_cfg["text_color"], loc="left")

        # Legend
        legend_handles = []
        legend_handles.append(
            plt.Line2D([0], [0], color=cfg.TODAY_LINE_COLOR,
                       linewidth=cfg.TODAY_LINE_WIDTH, linestyle=cfg.TODAY_LINE_STYLE,
                       label=f"Today ({today.strftime('%b %d, %Y')})")
        )
        for ref in ref_lines:
            legend_handles.append(
                plt.Line2D([0], [0], color=ref.color, linewidth=1.8,
                           linestyle=ref.style, label=ref.name)
            )
        for status, scolor in cfg.STATUS_COLORS.items():
            legend_handles.append(
                plt.Line2D([0], [0], marker="s", color="w", markerfacecolor=scolor,
                           markersize=8, label=status.replace("-", " ").title())
            )

        milestone_legend = [
            ("#27AE60", "Completed"),
            ("#F39C12", "At Risk (≤5 days)"),
            ("#E74C3C", "Overdue"),
            ("#5DADE2", "Upcoming"),
        ]
        for ms_color, ms_label in milestone_legend:
            legend_handles.append(
                plt.Line2D([0], [0], marker=cfg.MILESTONE_MARKER, color="w",
                           markerfacecolor=ms_color, markeredgecolor=cfg.MILESTONE_EDGE_COLOR,
                           markersize=10, label=ms_label)
            )
        ax.legend(handles=legend_handles, loc="upper right", fontsize=8,
                  framealpha=0.9, facecolor=theme_cfg["bg_color"],
                  labelcolor=theme_cfg["text_color"])

        fig.tight_layout()

        # Store figure reference
        self._current_figure = fig

        # Embed in PyQt6
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, self.chart_frame)

        self.chart_layout.addWidget(toolbar)
        self.chart_layout.addWidget(canvas)

        self.chart_canvas = canvas
        self.chart_toolbar = toolbar

        # Event handlers would need modification for PyQt - simplified version
        def on_milestone_pick(event):
            if event.mouseevent.button != 1:
                return
            for sc, proj_name, milestone in milestone_artists:
                if event.artist == sc:
                    dialog = _MilestoneTaskDialog(
                        self,
                        milestone_name=milestone.name,
                        milestone_id=milestone.milestone_id,
                        project_name=proj_name,
                        milestone_date=milestone.date,
                        can_edit=self.can_edit,
                        username=self.user["username"],
                        db_path=self.db_path
                    )
                    dialog.exec()
                    break

        canvas.mpl_connect('pick_event', on_milestone_pick)

    def _popout_timeline(self):
        projects, ref_lines = db.load_all(self.db_path)
        if not projects:
            QMessageBox.information(self, "Empty", "No projects in the database.")
            return
        render_timeline(
            projects=projects, today=datetime.date.today(),
            title="Project Timelines", reference_lines=ref_lines,
            theme=self.theme_var, show=True,
        )

    def _save_timeline_png(self):
        projects, ref_lines = db.load_all(self.db_path)
        if not projects:
            QMessageBox.information(self, "Empty", "No projects.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Chart", "",
            "PNG (*.png);;SVG (*.svg);;PDF (*.pdf)")
        if path:
            render_timeline(projects=projects, today=datetime.date.today(),
                            title="Project Timelines", reference_lines=ref_lines,
                            theme=self.theme_var, output_path=path, show=False)
            self._set_status(f"Saved to {path}")
            QMessageBox.information(self, "Saved", f"Chart saved to:\n{path}")

    def _save_timeline_html(self):
        projects, ref_lines = db.load_all(self.db_path)
        if not projects:
            QMessageBox.information(self, "Empty", "No projects.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save HTML", "", "HTML (*.html)")
        if path:
            render_timeline(projects=projects, today=datetime.date.today(),
                            title="Project Timelines", reference_lines=ref_lines,
                            theme=self.theme_var, html_path=path, show=False)
            self._set_status(f"HTML saved to {path}")
            QMessageBox.information(self, "Saved", f"Interactive chart saved to:\n{path}")

    def _export_pdf_report(self):
        if not REPORTLAB_AVAILABLE:
            QMessageBox.critical(
                self, "Missing Dependency",
                "reportlab is required for PDF export.\n\nInstall with:\npip install reportlab"
            )
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, "Save PDF Report",
            f"project_report_{datetime.date.today().strftime('%Y%m%d')}.pdf",
            "PDF files (*.pdf)"
        )

        if not filepath:
            return

        try:
            gantt_path = None
            if hasattr(self, '_current_figure') and self._current_figure:
                gantt_path = pathlib.Path(filepath).parent / "_temp_gantt.png"
                self._current_figure.savefig(str(gantt_path), dpi=150, bbox_inches='tight')

            projects, _ = db.load_all(self.db_path)

            generate_pdf_report(
                projects=projects, output_path=filepath,
                title="Project Status Report",
                include_gantt=True, include_milestones=True, include_kpis=True,
                gantt_image_path=gantt_path,
            )

            if gantt_path and gantt_path.exists():
                gantt_path.unlink()

            self._set_status(f"PDF report saved to {filepath}")
            QMessageBox.information(self, "Export Complete", f"PDF report saved successfully!\n\n{filepath}")

        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to generate PDF report:\n\n{e}")

    def _export_excel_report(self):
        if not OPENPYXL_AVAILABLE:
            QMessageBox.critical(
                self, "Missing Dependency",
                "openpyxl is required for Excel export.\n\nInstall with:\npip install openpyxl"
            )
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, "Save Excel Report",
            f"project_report_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            "Excel files (*.xlsx)"
        )

        if not filepath:
            return

        try:
            projects, _ = db.load_all(self.db_path)
            generate_excel_report(projects=projects, output_path=filepath, title="Project Status Report")
            self._set_status(f"Excel report saved to {filepath}")
            QMessageBox.information(self, "Export Complete", f"Excel report saved successfully!\n\n{filepath}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to generate Excel report:\n\n{e}")

    def _show_summary_dashboard(self):
        """Show summary dashboard dialog."""
        projects, ref_lines = db.load_all(self.db_path)
        today = datetime.date.today()

        if not projects:
            QMessageBox.information(self, "No Data", "No projects to summarize.")
            return

        dialog = _SummaryDashboardDialog(self, projects, today)
        dialog.exec()

    # ═════════════════════════════════════════════════════════════════════
    # TAB 2 — QCTP
    # ═════════════════════════════════════════════════════════════════════

    def _build_qctp_tab(self):
        layout = QVBoxLayout(self.tab_qctp)
        layout.setContentsMargins(10, 10, 10, 10)

        # Load QCTP template
        try:
            from timeline_tool.qctp_template import get_qctp_template
            self._qctp_template = get_qctp_template()
        except Exception as e:
            print(f"Warning: Could not load QCTP template: {e}")
            self._qctp_template = {}

        # Initialize current week
        self._current_week = datetime.date.today().isocalendar()[1]
        self._current_year = datetime.date.today().year
        self._qctp_projects_cache = []

        # Header
        header_layout = QHBoxLayout()
        header_label = QLabel("QCTP Overview")
        header_label.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        header_layout.addWidget(header_label)

        header_layout.addStretch()

        # Legend
        legend_layout = QHBoxLayout()
        legend_layout.addWidget(QLabel("Status: "))
        for text, color in [("● Green", "#16A34A"), ("● Orange", "#F97316"), ("● Red", "#DC2626")]:
            lbl = QLabel(text)
            lbl.setStyleSheet(f"color: {color}; font-weight: bold;")
            legend_layout.addWidget(lbl)
        header_layout.addLayout(legend_layout)
        layout.addLayout(header_layout)

        # Project selector
        selector_group = QGroupBox("  Select Project  ")
        selector_layout = QHBoxLayout(selector_group)

        selector_layout.addWidget(QLabel("Project:"))
        self.qctp_project_combo = QComboBox()
        self.qctp_project_combo.setMinimumWidth(300)
        self.qctp_project_combo.currentIndexChanged.connect(self._on_qctp_project_change)
        selector_layout.addWidget(self.qctp_project_combo)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_qctp_project_list)
        selector_layout.addWidget(refresh_btn)

        selector_layout.addStretch()

        if self.can_edit:
            save_btn = QPushButton("💾 Save All")
            save_btn.setObjectName("accentBtn")
            save_btn.clicked.connect(self._save_qctp)
            selector_layout.addWidget(save_btn)

        layout.addWidget(selector_group)

        # Phase Sub-Tabs
        self.qctp_notebook = QTabWidget()

        self.qctp_phases = [
            ("pre_program", "Pre-Program", "Before CM"),
            ("detailed_design", "Detailed Design Phase", "CM to Sync 5/SHRM"),
            ("industrialization", "Industrialization Phase", "Sync5 to SOP"),
        ]

        self.qctp_categories = [
            ("quality", "Quality", "🎯"),
            ("cost", "Cost", "💰"),
            ("time", "Time", "⏱️"),
            ("performance", "Performance", "📈"),
        ]

        self.qctp_line_widgets = {}
        self.qctp_notes_widgets = {}
        self._week_labels = []

        for phase_key, phase_name, phase_desc in self.qctp_phases:
            phase_tab = QWidget()
            phase_layout = QVBoxLayout(phase_tab)
            phase_layout.setContentsMargins(5, 5, 5, 5)

            # Phase description with week navigation
            desc_layout = QHBoxLayout()

            left_desc = QHBoxLayout()
            phase_title = QLabel(f"📋 {phase_name}")
            phase_title.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
            left_desc.addWidget(phase_title)
            phase_desc_lbl = QLabel(f"  ({phase_desc})")
            phase_desc_lbl.setStyleSheet("color: #666666; font-style: italic;")
            left_desc.addWidget(phase_desc_lbl)
            left_desc.addStretch()
            desc_layout.addLayout(left_desc)

            # Week navigation
            week_layout = QHBoxLayout()
            prev_btn = QPushButton("◀ Previous Week")
            prev_btn.clicked.connect(lambda: self._change_week(-1))
            week_layout.addWidget(prev_btn)

            week_label = QLabel(f"Week {self._current_week}")
            week_label.setStyleSheet("background-color: #0067C0; color: white; padding: 5px 15px; font-weight: bold;")
            week_layout.addWidget(week_label)
            self._week_labels.append(week_label)

            next_btn = QPushButton("Next Week ▶")
            next_btn.clicked.connect(lambda: self._change_week(1))
            week_layout.addWidget(next_btn)

            desc_layout.addLayout(week_layout)
            phase_layout.addLayout(desc_layout)

            # Main content - horizontal splitter
            splitter = QSplitter(Qt.Orientation.Horizontal)

            # Left: QCTP grid in scroll area
            left_scroll = QScrollArea()
            left_scroll.setWidgetResizable(True)
            left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

            left_content = QWidget()
            left_layout = QGridLayout(left_content)
            left_layout.setSpacing(10)

            self.qctp_line_widgets[phase_key] = {}

            positions = [(0, 0), (0, 1), (1, 0), (1, 1)]
            for (cat_key, cat_name, cat_icon), (row, col) in zip(self.qctp_categories, positions):
                category_widget = self._create_qctp_category_box(phase_key, cat_key, cat_name, cat_icon)
                left_layout.addWidget(category_widget, row, col)

            left_scroll.setWidget(left_content)
            splitter.addWidget(left_scroll)

            # Right: Notes panels
            right_widget = QWidget()
            right_widget.setFixedWidth(400)
            right_layout = QVBoxLayout(right_widget)

            self.qctp_notes_widgets[phase_key] = {}

            # Highlights
            highlights_group = QGroupBox("HIGHLIGHTS (Focus on main difficulties all phases)")
            highlights_layout = QVBoxLayout(highlights_group)
            highlights_text = QTextEdit()
            highlights_text.setReadOnly(not self.can_edit)
            highlights_layout.addWidget(highlights_text)
            right_layout.addWidget(highlights_group)
            self.qctp_notes_widgets[phase_key]["highlights"] = highlights_text

            # Red Points
            red_group = QGroupBox("RED points explanation")
            red_group.setStyleSheet("QGroupBox::title { color: #DC2626; }")
            red_layout = QVBoxLayout(red_group)
            red_text = QTextEdit()
            red_text.setReadOnly(not self.can_edit)
            red_layout.addWidget(red_text)
            right_layout.addWidget(red_group)
            self.qctp_notes_widgets[phase_key]["red_points"] = red_text

            # VSSQ Escalation
            escalation_group = QGroupBox("VSSQ Escalation")
            escalation_layout = QVBoxLayout(escalation_group)
            escalation_text = QTextEdit()
            escalation_text.setReadOnly(not self.can_edit)
            escalation_layout.addWidget(escalation_text)
            right_layout.addWidget(escalation_group)
            self.qctp_notes_widgets[phase_key]["escalation"] = escalation_text

            splitter.addWidget(right_widget)
            splitter.setSizes([600, 400])

            phase_layout.addWidget(splitter)
            self.qctp_notebook.addTab(phase_tab, f"  {phase_name}  ")

        layout.addWidget(self.qctp_notebook)
        self._refresh_qctp_project_list()

    def _create_qctp_category_box(self, phase_key: str, cat_key: str, cat_name: str, cat_icon: str) -> QGroupBox:
        """Create a QCTP category box with line items."""
        all_descriptions = self._qctp_template.get(phase_key, {}).get(cat_key, [])
        descriptions_with_index = [
            (idx + 1, desc.strip())
            for idx, desc in enumerate(all_descriptions)
            if desc and desc.strip()
        ]

        group = QGroupBox(f"{cat_icon} {cat_name}")
        layout = QVBoxLayout(group)
        layout.setSpacing(5)

        self.qctp_line_widgets[phase_key][cat_key] = []

        if not descriptions_with_index:
            no_data = QLabel("No items defined")
            no_data.setStyleSheet("color: #999999; font-style: italic;")
            no_data.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(no_data)
            return group

        # Header row
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("#"), stretch=1)
        header_layout.addWidget(QLabel("Description"), stretch=10)
        header_layout.addWidget(QLabel("Status"), stretch=2)
        header_layout.addWidget(QLabel("Remarks"), stretch=4)
        header_layout.addWidget(QLabel("Attach"), stretch=1)
        layout.addLayout(header_layout)

        # Line items
        for display_num, (original_line_num, description) in enumerate(descriptions_with_index, 1):
            line_widgets = self._create_qctp_line_item(phase_key, cat_key, display_num, original_line_num, description)
            self.qctp_line_widgets[phase_key][cat_key].append(line_widgets)

            row_layout = QHBoxLayout()
            row_layout.addWidget(QLabel(str(display_num)), stretch=1)

            desc_label = QLabel(description[:55] + "..." if len(description) > 55 else description)
            desc_label.setStyleSheet("background-color: #F0F0F0; padding: 3px;")
            desc_label.setToolTip(description)
            row_layout.addWidget(desc_label, stretch=10)

            row_layout.addWidget(line_widgets["status_combo"], stretch=2)
            row_layout.addWidget(line_widgets["remarks_entry"], stretch=4)
            row_layout.addWidget(line_widgets["attach_btn"], stretch=1)

            layout.addLayout(row_layout)

        return group

    def _create_qctp_line_item(self, phase_key: str, cat_key: str, display_num: int,
                               original_line_num: int, description: str) -> dict:
        """Create a single QCTP line item row."""
        status_combo = QComboBox()
        status_combo.addItems(["", "Green", "Orange", "Red"])
        status_combo.setEnabled(self.can_edit)

        remarks_entry = QLineEdit()
        remarks_entry.setReadOnly(not self.can_edit)

        attach_btn = QPushButton("📎")
        attach_btn.setFixedWidth(30)

        attachment_path = ""

        def browse_attachment():
            nonlocal attachment_path
            filepath, _ = QFileDialog.getOpenFileName(
                self, f"Select attachment for {cat_key.title()} Item {display_num}",
                "", "All files (*.*)"
            )
            if filepath:
                attachment_path = filepath
                attach_btn.setText("✓")
                attach_btn.setStyleSheet("color: #16A34A;")

        attach_btn.clicked.connect(browse_attachment)

        return {
            "display_num": display_num,
            "line_num": original_line_num,
            "description": description,
            "status_combo": status_combo,
            "remarks_entry": remarks_entry,
            "attach_btn": attach_btn,
            "get_attachment": lambda: attachment_path,
            "set_attachment": lambda p: self._set_qctp_attachment(attach_btn, p),
        }

    def _set_qctp_attachment(self, btn, path):
        if path:
            btn.setText("✓")
            btn.setStyleSheet("color: #16A34A;")
        else:
            btn.setText("📎")
            btn.setStyleSheet("")

    def _change_week(self, delta: int):
        self._current_week += delta
        if self._current_week < 1:
            self._current_year -= 1
            self._current_week = 52
        elif self._current_week > 52:
            self._current_year += 1
            self._current_week = 1

        for label in self._week_labels:
            label.setText(f"Week {self._current_week}")

        self._on_qctp_project_change()
        self._set_status(f"Switched to Week {self._current_week}, {self._current_year}")

    def _get_qctp_selected_project_id(self) -> int | None:
        idx = self.qctp_project_combo.currentIndex()
        if idx < 0 or not self._qctp_projects_cache:
            return None
        return self._qctp_projects_cache[idx]["id"]

    def _refresh_qctp_project_list(self):
        with db._connect(self.db_path) as conn:
            rows = conn.execute("SELECT id, name FROM projects ORDER BY start_date").fetchall()
        self._qctp_projects_cache = list(rows)
        self.qctp_project_combo.clear()
        self.qctp_project_combo.addItems([f"{r['id']} — {r['name']}" for r in rows])
        if rows:
            self.qctp_project_combo.setCurrentIndex(0)
            self._on_qctp_project_change()

    def _on_qctp_project_change(self, index=None):
        pid = self._get_qctp_selected_project_id()
        if pid is None:
            return

        for phase_key, _, _ in self.qctp_phases:
            for cat_key, _, _ in self.qctp_categories:
                try:
                    line_items = db.get_qctp_line_items(pid, phase_key, cat_key, self.db_path)
                except Exception:
                    line_items = []

                items_by_line = {item["line_number"]: item for item in line_items}

                for widget_data in self.qctp_line_widgets[phase_key][cat_key]:
                    line_num = widget_data["line_num"]
                    item = items_by_line.get(line_num, {})

                    saved_status = item.get("status", "")
                    widget_data["status_combo"].setCurrentText(saved_status)
                    widget_data["remarks_entry"].setText(item.get("remarks", ""))

                    attachment = item.get("attachment_path", "")
                    widget_data["set_attachment"](attachment)

            # Load notes
            if phase_key in self.qctp_notes_widgets:
                try:
                    notes = db.get_qctp_notes(pid, phase_key, self._current_week, self._current_year, self.db_path)
                except Exception:
                    notes = {}

                for note_key in ["highlights", "red_points", "escalation"]:
                    text_widget = self.qctp_notes_widgets[phase_key].get(note_key)
                    if text_widget:
                        text_widget.setPlainText(notes.get(note_key, ""))

    def _save_qctp(self):
        pid = self._get_qctp_selected_project_id()
        if pid is None:
            QMessageBox.warning(self, "Select", "Select a project first.")
            return

        try:
            saved_count = 0

            for phase_key, _, _ in self.qctp_phases:
                for cat_key, _, _ in self.qctp_categories:
                    for widget_data in self.qctp_line_widgets[phase_key][cat_key]:
                        line_num = widget_data["line_num"]
                        description = widget_data["description"]
                        status = widget_data["status_combo"].currentText()
                        remarks = widget_data["remarks_entry"].text().strip()
                        attachment = widget_data["get_attachment"]()

                        db.save_qctp_line_item(
                            project_id=pid, phase=phase_key, category=cat_key,
                            line_number=line_num, description=description,
                            status=status, remarks=remarks, attachment_path=attachment,
                            username=self.user["username"], db_path=self.db_path
                        )
                        saved_count += 1

                # Save notes
                if phase_key in self.qctp_notes_widgets:
                    notes_data = {}
                    for note_key in ["highlights", "red_points", "escalation"]:
                        text_widget = self.qctp_notes_widgets[phase_key].get(note_key)
                        if text_widget:
                            notes_data[note_key] = text_widget.toPlainText().strip()

                    db.save_qctp_notes(
                        project_id=pid, phase=phase_key,
                        week_number=self._current_week, year=self._current_year,
                        highlights=notes_data.get("highlights", ""),
                        red_points=notes_data.get("red_points", ""),
                        escalation=notes_data.get("escalation", ""),
                        username=self.user["username"], db_path=self.db_path
                    )

            with db._connect(self.db_path) as conn:
                db.log_action(conn, self.user["username"], "SAVE_QCTP",
                              f"Saved {saved_count} QCTP line items for project ID {pid}, Week {self._current_week}")

            self._set_status(f"QCTP saved successfully ({saved_count} items, Week {self._current_week})")
            QMessageBox.information(self, "Saved", f"QCTP data saved successfully!\n\n{saved_count} line items saved for Week {self._current_week}.")

        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save QCTP data:\n\n{str(e)}")

    # ═════════════════════════════════════════════════════════════════════
    # TAB 3 — ACTIVITIES
    # ═════════════════════════════════════════════════════════════════════

    def _build_activities_tab(self):
        layout = QVBoxLayout(self.tab_activities)
        layout.setContentsMargins(10, 10, 10, 10)

        # Header
        header_label = QLabel("Weekly Activities")
        header_label.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        layout.addWidget(header_label)

        # Project + Week Selector
        selector_group = QGroupBox("  Select Project & Week  ")
        selector_layout = QHBoxLayout(selector_group)

        selector_layout.addWidget(QLabel("Project:"))
        self.act_project_combo = QComboBox()
        self.act_project_combo.setMinimumWidth(250)
        self.act_project_combo.currentIndexChanged.connect(self._on_activity_project_change)
        selector_layout.addWidget(self.act_project_combo)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_activity_project_list)
        selector_layout.addWidget(refresh_btn)

        # Week navigation
        self._act_current_week = datetime.date.today().isocalendar()[1]
        self._act_current_year = datetime.date.today().year

        prev_btn = QPushButton("◀")
        prev_btn.clicked.connect(lambda: self._change_activity_week(-1))
        selector_layout.addWidget(prev_btn)

        self.act_week_label = QLabel(f"Week {self._act_current_week}, {self._act_current_year}")
        self.act_week_label.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        selector_layout.addWidget(self.act_week_label)

        next_btn = QPushButton("▶")
        next_btn.clicked.connect(lambda: self._change_activity_week(1))
        selector_layout.addWidget(next_btn)

        selector_layout.addStretch()

        if self.can_edit:
            add_btn = QPushButton("➕ Add Activity")
            add_btn.setObjectName("accentBtn")
            add_btn.clicked.connect(self._add_activity)
            selector_layout.addWidget(add_btn)

            edit_btn = QPushButton("✏️ Edit")
            edit_btn.clicked.connect(self._edit_activity)
            selector_layout.addWidget(edit_btn)

            del_btn = QPushButton("🗑️ Delete")
            del_btn.clicked.connect(self._delete_activity)
            selector_layout.addWidget(del_btn)

        layout.addWidget(selector_group)

        # Activities Table
        table_group = QGroupBox("  Activities  ")
        table_layout = QVBoxLayout(table_group)

        self.act_table = QTableWidget()
        self.act_table.setColumnCount(11)
        self.act_table.setHorizontalHeaderLabels([
            "ID", "Project", "Activity", "Start Date", "End Date", "Time Taken",
            "Members", "Hard Points", "Status", "Attachment", "Updated By"
        ])
        self.act_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.act_table.setAlternatingRowColors(True)
        self.act_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.act_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table_layout.addWidget(self.act_table)

        layout.addWidget(table_group)

        self._act_projects_cache = []
        self._refresh_activity_project_list()

    def _change_activity_week(self, delta: int):
        self._act_current_week += delta
        if self._act_current_week > 52:
            self._act_current_week = 1
            self._act_current_year += 1
        elif self._act_current_week < 1:
            self._act_current_week = 52
            self._act_current_year -= 1
        self.act_week_label.setText(f"Week {self._act_current_week}, {self._act_current_year}")
        self._refresh_activities_table()

    def _refresh_activity_project_list(self):
        with db._connect(self.db_path) as conn:
            rows = conn.execute("SELECT id, name FROM projects ORDER BY start_date").fetchall()
        self._act_projects_cache = list(rows)

        self.act_project_combo.clear()
        self.act_project_combo.addItem("All Projects")
        for r in rows:
            self.act_project_combo.addItem(f"{r['name']} ({r['id']})")

        self.act_project_combo.setCurrentIndex(0)
        self._on_activity_project_change()

    def _get_act_selected_project_id(self):
        idx = self.act_project_combo.currentIndex()
        if idx == 0:
            return None
        return self._act_projects_cache[idx - 1]["id"]

    def _on_activity_project_change(self, index=None):
        self._refresh_activities_table()

    def _refresh_activities_table(self):
        self.act_table.setRowCount(0)

        project_id = self._get_act_selected_project_id()

        if project_id is None:
            activities = []
            for proj in self._act_projects_cache:
                activities += db.get_activities(
                    proj["id"], self._act_current_week, self._act_current_year, db_path=self.db_path
                )
        else:
            activities = db.get_activities(
                project_id, self._act_current_week, self._act_current_year, db_path=self.db_path
            )

        for act in activities:
            row = self.act_table.rowCount()
            self.act_table.insertRow(row)

            attachment_display = (
                pathlib.Path(act["attachment_path"]).name if act["attachment_path"] else ""
            )

            values = [
                str(act["id"]), act["project_name"], act["activity_name"],
                act["start_date"], act["end_date"], act["time_taken"],
                act["members"], act["hard_points"], act["status"],
                attachment_display, act.get("updated_by", "")
            ]

            for col, val in enumerate(values):
                item = QTableWidgetItem(val)
                self.act_table.setItem(row, col, item)

            # Color by status
            if act["status"] == "Completed":
                for col in range(self.act_table.columnCount()):
                    self.act_table.item(row, col).setBackground(QColor("#DCFCE7"))
            elif act["status"] == "WIP":
                for col in range(self.act_table.columnCount()):
                    self.act_table.item(row, col).setBackground(QColor("#FEF9C3"))

    def _get_selected_activity_id(self) -> int | None:
        row = self.act_table.currentRow()
        if row < 0:
            return None
        return int(self.act_table.item(row, 0).text())

    def _add_activity(self):
        project_id = self._get_act_selected_project_id()
        if project_id is None:
            QMessageBox.warning(self, "No Project", "Please select a project first.")
            return

        dialog = _ActivityDialog(self, "Add Activity", self.db_path,
                                 week=self._act_current_week, year=self._act_current_year)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            db.add_activity(
                project_id=project_id,
                week_number=self._act_current_week,
                year=self._act_current_year,
                username=self.user["username"],
                db_path=self.db_path,
                **dialog.result
            )
            self._refresh_activities_table()
            self._set_status("Activity added successfully")

    def _edit_activity(self):
        activity_id = self._get_selected_activity_id()
        if activity_id is None:
            QMessageBox.warning(self, "No Selection", "Please select an activity to edit.")
            return

        with db._connect(self.db_path) as conn:
            row = conn.execute("SELECT * FROM activities WHERE id = ?", (activity_id,)).fetchone()
        if not row:
            return

        initial = {
            "activity_name": row["activity_name"],
            "start_date": row["start_date"],
            "end_date": row["end_date"],
            "time_taken": row["time_taken"],
            "members": row["members"],
            "hard_points": row["hard_points"],
            "status": row["status"],
            "attachment_path": row["attachment_path"],
        }

        dialog = _ActivityDialog(self, "Edit Activity", self.db_path,
                                 week=self._act_current_week, year=self._act_current_year,
                                 initial=initial)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            db.update_activity(
                activity_id=activity_id,
                username=self.user["username"],
                db_path=self.db_path,
                **dialog.result
            )
            self._refresh_activities_table()
            self._set_status("Activity updated successfully")

    def _delete_activity(self):
        activity_id = self._get_selected_activity_id()
        if activity_id is None:
            QMessageBox.warning(self, "No Selection", "Please select an activity to delete.")
            return

        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Are you sure you want to delete this activity?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            db.delete_activity(activity_id, username=self.user["username"], db_path=self.db_path)
            self._refresh_activities_table()
            self._set_status("Activity deleted")

    # ═════════════════════════════════════════════════════════════════════
    # TAB 4 — PROJECTS
    # ═════════════════════════════════════════════════════════════════════

    def _build_projects_tab(self):
        layout = QVBoxLayout(self.tab_projects)
        layout.setContentsMargins(10, 10, 10, 10)

        # Header
        header_layout = QHBoxLayout()
        header_label = QLabel("Project Manager")
        header_label.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        header_layout.addWidget(header_label)

        header_layout.addStretch()

        if self.can_edit:
            add_btn = QPushButton("➕ Add Project")
            add_btn.setObjectName("accentBtn")
            add_btn.clicked.connect(self._add_project)
            header_layout.addWidget(add_btn)

            edit_btn = QPushButton("✏️ Edit")
            edit_btn.clicked.connect(self._edit_project)
            header_layout.addWidget(edit_btn)

            del_btn = QPushButton("🗑️ Delete")
            del_btn.clicked.connect(self._delete_project)
            header_layout.addWidget(del_btn)

            import_btn = QPushButton("📥 Import JSON")
            import_btn.clicked.connect(self._import_json)
            header_layout.addWidget(import_btn)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_projects_tab)
        header_layout.addWidget(refresh_btn)

        layout.addLayout(header_layout)

        # Table
        table_group = QGroupBox("  All Projects  ")
        table_layout = QVBoxLayout(table_group)

        self.proj_table = QTableWidget()
        self.proj_table.setColumnCount(10)
        self.proj_table.setHorizontalHeaderLabels([
            "ID", "Name", "Start Date", "End Date", "Status",
            "Dev Region", "Sales Region", "Color", "Created By", "Last Updated"
        ])
        self.proj_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.proj_table.setAlternatingRowColors(True)
        self.proj_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.proj_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table_layout.addWidget(self.proj_table)

        layout.addWidget(table_group)

        self._refresh_projects_tab()

    def _refresh_projects_tab(self):
        self.proj_table.setRowCount(0)
        today = datetime.date.today()

        with db._connect(self.db_path) as conn:
            rows = conn.execute("SELECT * FROM projects ORDER BY start_date").fetchall()

            for r in rows:
                row = self.proj_table.rowCount()
                self.proj_table.insertRow(row)

                dev_reg = r["dev_region"] if "dev_region" in r.keys() else ""
                sales_reg = r["sales_region"] if "sales_region" in r.keys() else ""

                # Compute status
                from timeline_tool.models import Milestone as MilestoneModel
                ms_rows = conn.execute(
                    "SELECT id, name, date FROM milestones WHERE project_id = ? ORDER BY date",
                    (r["id"],)
                ).fetchall()
                temp_milestones = []
                for ms_row in ms_rows:
                    task_statuses = {}
                    task_rows = conn.execute(
                        "SELECT status, COUNT(*) as cnt FROM milestone_tasks WHERE milestone_id = ? GROUP BY status",
                        (ms_row["id"],)
                    ).fetchall()
                    for tr in task_rows:
                        task_statuses[tr["status"]] = tr["cnt"]
                    temp_milestones.append(MilestoneModel(
                        name=ms_row["name"],
                        date=datetime.date.fromisoformat(ms_row["date"]),
                        milestone_id=ms_row["id"],
                        task_statuses=task_statuses,
                    ))

                temp_proj = Project(
                    name=r["name"],
                    start_date=datetime.date.fromisoformat(r["start_date"]),
                    end_date=datetime.date.fromisoformat(r["end_date"]),
                    milestones=temp_milestones,
                )
                computed = temp_proj.computed_status(today).replace("-", " ").title()

                values = [
                    str(r["id"]), r["name"], r["start_date"], r["end_date"],
                    computed, dev_reg, sales_reg,
                    r["color"] or "auto", r["created_by"] or "", r["updated_at"],
                ]

                for col, val in enumerate(values):
                    item = QTableWidgetItem(val)
                    self.proj_table.setItem(row, col, item)

        self._set_status(f"Projects tab refreshed — {self.proj_table.rowCount()} project(s)")

    def _get_selected_proj_id(self) -> int | None:
        row = self.proj_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Please select a project first.")
            return None
        return int(self.proj_table.item(row, 0).text())

    def _add_project(self):
        dialog = _ProjectDialog(self, "Add Project")
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                db.add_project(**dialog.result, username=self.user["username"], db_path=self.db_path)
                self._refresh_projects_tab()
                self._refresh_dashboard()
                self._set_status(f"Project '{dialog.result['name']}' added")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _edit_project(self):
        pid = self._get_selected_proj_id()
        if pid is None:
            return

        with db._connect(self.db_path) as conn:
            proj = conn.execute("SELECT * FROM projects WHERE id = ?", (pid,)).fetchone()

        initial = {
            "name": proj["name"], "start_date": proj["start_date"],
            "end_date": proj["end_date"], "color": proj["color"] or "",
            "status": proj["status"],
            "dev_region": proj["dev_region"] if "dev_region" in proj.keys() else "",
            "sales_region": proj["sales_region"] if "sales_region" in proj.keys() else "",
        }

        dialog = _ProjectDialog(self, "Edit Project", initial)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                db.update_project(pid, **dialog.result, username=self.user["username"], db_path=self.db_path)
                self._refresh_projects_tab()
                self._refresh_dashboard()
                self._set_status("Project updated")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _delete_project(self):
        pid = self._get_selected_proj_id()
        if pid is None:
            return

        reply = QMessageBox.question(
            self, "Confirm",
            "Delete this project and all its milestones/phases?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            db.delete_project(pid, username=self.user["username"], db_path=self.db_path)
            self._refresh_projects_tab()
            self._refresh_dashboard()
            self._set_status("Project deleted")

    def _import_json(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select JSON file", "",
            "JSON files (*.json);;All files (*.*)"
        )
        if path:
            try:
                db.import_from_json(path, username=self.user["username"], db_path=self.db_path)
                self._refresh_projects_tab()
                self._refresh_dashboard()
                self._set_status("JSON data imported successfully")
                QMessageBox.information(self, "Success", "Data imported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    # ═════════════════════════════════════════════════════════════════════
    # TAB 5 — MILESTONES & PHASES
    # ═════════════════════════════════════════════════════════════════════

    def _build_milestones_tab(self):
        layout = QVBoxLayout(self.tab_milestones)
        layout.setContentsMargins(10, 10, 10, 10)

        # Header
        header_label = QLabel("Milestones & Phases")
        header_label.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        layout.addWidget(header_label)

        # Project selector
        selector_group = QGroupBox("  Select Project  ")
        selector_layout = QHBoxLayout(selector_group)

        selector_layout.addWidget(QLabel("Project:"))
        self.ms_project_combo = QComboBox()
        self.ms_project_combo.setMinimumWidth(300)
        self.ms_project_combo.currentIndexChanged.connect(self._on_ms_project_change)
        selector_layout.addWidget(self.ms_project_combo)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_ms_project_list)
        selector_layout.addWidget(refresh_btn)
        selector_layout.addStretch()

        layout.addWidget(selector_group)

        # Splitter for milestones and phases
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Milestones panel
        ms_group = QGroupBox("  🎯 Milestones  ")
        ms_layout = QVBoxLayout(ms_group)

        self.ms_table = QTableWidget()
        self.ms_table.setColumnCount(4)
        self.ms_table.setHorizontalHeaderLabels(["ID", "Name", "Date", "Updated By"])
        self.ms_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.ms_table.setAlternatingRowColors(True)
        self.ms_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.ms_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        ms_layout.addWidget(self.ms_table)

        if self.can_edit:
            ms_btn_layout = QHBoxLayout()
            add_ms_btn = QPushButton("➕ Add")
            add_ms_btn.setObjectName("accentBtn")
            add_ms_btn.clicked.connect(self._add_milestone)
            ms_btn_layout.addWidget(add_ms_btn)

            edit_ms_btn = QPushButton("✏️ Edit")
            edit_ms_btn.clicked.connect(self._edit_milestone)
            ms_btn_layout.addWidget(edit_ms_btn)

            del_ms_btn = QPushButton("🗑️ Delete")
            del_ms_btn.clicked.connect(self._delete_milestone)
            ms_btn_layout.addWidget(del_ms_btn)

            ms_layout.addLayout(ms_btn_layout)

        splitter.addWidget(ms_group)

        # Phases panel
        ph_group = QGroupBox("  📐 Phases  ")
        ph_layout = QVBoxLayout(ph_group)

        self.ph_table = QTableWidget()
        self.ph_table.setColumnCount(4)
        self.ph_table.setHorizontalHeaderLabels(["ID", "Name", "Start", "End"])
        self.ph_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.ph_table.setAlternatingRowColors(True)
        self.ph_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.ph_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        ph_layout.addWidget(self.ph_table)

        if self.can_edit:
            ph_btn_layout = QHBoxLayout()
            add_ph_btn = QPushButton("➕ Add")
            add_ph_btn.setObjectName("accentBtn")
            add_ph_btn.clicked.connect(self._add_phase)
            ph_btn_layout.addWidget(add_ph_btn)

            edit_ph_btn = QPushButton("✏️ Edit")
            edit_ph_btn.clicked.connect(self._edit_phase)
            ph_btn_layout.addWidget(edit_ph_btn)

            del_ph_btn = QPushButton("🗑️ Delete")
            del_ph_btn.clicked.connect(self._delete_phase)
            ph_btn_layout.addWidget(del_ph_btn)

            ph_layout.addLayout(ph_btn_layout)

        splitter.addWidget(ph_group)
        layout.addWidget(splitter)

        self._ms_projects_cache = []
        self._refresh_ms_project_list()

    def _refresh_ms_project_list(self):
        with db._connect(self.db_path) as conn:
            rows = conn.execute("SELECT id, name FROM projects ORDER BY start_date").fetchall()
        self._ms_projects_cache = list(rows)
        self.ms_project_combo.clear()
        self.ms_project_combo.addItems([f"{r['id']} — {r['name']}" for r in rows])
        if rows:
            self.ms_project_combo.setCurrentIndex(0)
            self._on_ms_project_change()

    def _get_ms_selected_project_id(self) -> int | None:
        idx = self.ms_project_combo.currentIndex()
        if idx < 0 or not self._ms_projects_cache:
            return None
        return self._ms_projects_cache[idx]["id"]

    def _on_ms_project_change(self, index=None):
        pid = self._get_ms_selected_project_id()
        if pid is None:
            return

        with db._connect(self.db_path) as conn:
            ms_rows = conn.execute(
                "SELECT id, name, date, created_by FROM milestones WHERE project_id = ? ORDER BY date",
                (pid,)
            ).fetchall()
            ph_rows = conn.execute(
                "SELECT id, name, start_date, end_date FROM phases WHERE project_id = ? ORDER BY start_date",
                (pid,)
            ).fetchall()

        self.ms_table.setRowCount(0)
        for m in ms_rows:
            row = self.ms_table.rowCount()
            self.ms_table.insertRow(row)
            for col, val in enumerate([str(m["id"]), m["name"], m["date"], m["created_by"] or ""]):
                self.ms_table.setItem(row, col, QTableWidgetItem(val))

        self.ph_table.setRowCount(0)
        for p in ph_rows:
            row = self.ph_table.rowCount()
            self.ph_table.insertRow(row)
            for col, val in enumerate([str(p["id"]), p["name"], p["start_date"], p["end_date"]]):
                self.ph_table.setItem(row, col, QTableWidgetItem(val))

    def _add_milestone(self):
        pid = self._get_ms_selected_project_id()
        if pid is None:
            QMessageBox.warning(self, "Select", "Select a project first.")
            return

        dialog = _MilestoneDialog(self, "Add Milestone")
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                db.add_milestone(pid, **dialog.result, username=self.user["username"], db_path=self.db_path)
                self._on_ms_project_change()
                self._set_status("Milestone added")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _edit_milestone(self):
        row = self.ms_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Select a milestone first.")
            return

        ms_id = int(self.ms_table.item(row, 0).text())
        initial = {
            "name": self.ms_table.item(row, 1).text(),
            "date": self.ms_table.item(row, 2).text()
        }

        dialog = _MilestoneDialog(self, "Edit Milestone", initial)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                db.update_milestone(ms_id, **dialog.result, username=self.user["username"], db_path=self.db_path)
                self._on_ms_project_change()
                self._set_status("Milestone updated")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _delete_milestone(self):
        row = self.ms_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Select a milestone first.")
            return

        ms_id = int(self.ms_table.item(row, 0).text())
        reply = QMessageBox.question(
            self, "Confirm", "Delete this milestone?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            db.delete_milestone(ms_id, username=self.user["username"], db_path=self.db_path)
            self._on_ms_project_change()
            self._set_status("Milestone deleted")

    def _add_phase(self):
        pid = self._get_ms_selected_project_id()
        if pid is None:
            QMessageBox.warning(self, "Select", "Select a project first.")
            return

        dialog = _PhaseDialog(self, "Add Phase")
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                db.add_phase(pid, **dialog.result, username=self.user["username"], db_path=self.db_path)
                self._on_ms_project_change()
                self._set_status("Phase added")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _edit_phase(self):
        row = self.ph_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Select a phase first.")
            return

        ph_id = int(self.ph_table.item(row, 0).text())
        initial = {
            "name": self.ph_table.item(row, 1).text(),
            "start_date": self.ph_table.item(row, 2).text(),
            "end_date": self.ph_table.item(row, 3).text()
        }

        dialog = _PhaseDialog(self, "Edit Phase", initial)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                db.update_phase(ph_id, **dialog.result, username=self.user["username"], db_path=self.db_path)
                self._on_ms_project_change()
                self._set_status("Phase updated")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _delete_phase(self):
        row = self.ph_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Select a phase first.")
            return

        ph_id = int(self.ph_table.item(row, 0).text())
        reply = QMessageBox.question(
            self, "Confirm", "Delete this phase?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            db.delete_phase(ph_id, username=self.user["username"], db_path=self.db_path)
            self._on_ms_project_change()
            self._set_status("Phase deleted")

    # ═════════════════════════════════════════════════════════════════════
    # TAB 6 — RESOURCES
    # ═════════════════════════════════════════════════════════════════════

    def _build_resources_tab(self):
        layout = QVBoxLayout(self.tab_resources)
        layout.setContentsMargins(10, 10, 10, 10)

        # Header
        header_label = QLabel("Resource & Team Management")
        header_label.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        layout.addWidget(header_label)

        # Sub-notebook for Resources
        res_notebook = QTabWidget()

        # Subtab 1: Team Members
        team_tab = QWidget()
        res_notebook.addTab(team_tab, "  👤 Team Members  ")
        self._build_team_subtab(team_tab)

        # Subtab 2: Assignments
        assign_tab = QWidget()
        res_notebook.addTab(assign_tab, "  📋 Assignments  ")
        self._build_assignments_subtab(assign_tab)

        # Subtab 3: Workload
        workload_tab = QWidget()
        res_notebook.addTab(workload_tab, "  📊 Workload  ")
        self._build_workload_subtab(workload_tab)

        layout.addWidget(res_notebook)

    def _build_team_subtab(self, tab):
        layout = QVBoxLayout(tab)

        # Buttons
        btn_layout = QHBoxLayout()

        if self.can_edit:
            add_btn = QPushButton("➕ Add Team Member")
            add_btn.setObjectName("accentBtn")
            add_btn.clicked.connect(self._add_team_member)
            btn_layout.addWidget(add_btn)

            edit_btn = QPushButton("✏️ Edit")
            edit_btn.clicked.connect(self._edit_team_member)
            btn_layout.addWidget(edit_btn)

            del_btn = QPushButton("🗑️ Delete")
            del_btn.clicked.connect(self._delete_team_member)
            btn_layout.addWidget(del_btn)

        if self.can_manage:
            import_btn = QPushButton("📥 Import from Excel")
            import_btn.clicked.connect(self._import_team_from_excel)
            btn_layout.addWidget(import_btn)

        btn_layout.addStretch()

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_team_list)
        btn_layout.addWidget(refresh_btn)

        layout.addLayout(btn_layout)

        # Team table
        self.team_table = QTableWidget()
        self.team_table.setColumnCount(6)
        self.team_table.setHorizontalHeaderLabels(["ID", "Name", "Role", "Department", "Email", "Capacity %"])
        self.team_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.team_table.setAlternatingRowColors(True)
        self.team_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.team_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.team_table)

        self._refresh_team_list()

    def _refresh_team_list(self):
        self.team_table.setRowCount(0)
        try:
            resources = get_all_resources(self.db_path)
            for res in resources:
                row = self.team_table.rowCount()
                self.team_table.insertRow(row)
                values = [str(res.id), res.name, res.role, res.department, res.email, f"{res.allocation_pct:.0f}%"]
                for col, val in enumerate(values):
                    self.team_table.setItem(row, col, QTableWidgetItem(val))
        except Exception as e:
            print(f"Error loading resources: {e}")

    def _add_team_member(self):
        dialog = _ResourceDialog(self, "Add Team Member")
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                add_resource(**dialog.result, username=self.user["username"], db_path=self.db_path)
                self._refresh_team_list()
                self._set_status(f"Team member '{dialog.result['name']}' added")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to add team member:\n{e}")

    def _edit_team_member(self):
        row = self.team_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Please select a team member to edit.")
            return

        resource_id = int(self.team_table.item(row, 0).text())
        initial = {
            "name": self.team_table.item(row, 1).text(),
            "role": self.team_table.item(row, 2).text(),
            "department": self.team_table.item(row, 3).text(),
            "email": self.team_table.item(row, 4).text(),
            "allocation_pct": float(self.team_table.item(row, 5).text().replace("%", "")),
        }

        dialog = _ResourceDialog(self, "Edit Team Member", initial)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                update_resource(resource_id, **dialog.result,
                               username=self.user["username"], db_path=self.db_path)
                self._refresh_team_list()
                self._set_status("Team member updated")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to update:\n{e}")

    def _delete_team_member(self):
        row = self.team_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Please select a team member to delete.")
            return

        resource_id = int(self.team_table.item(row, 0).text())
        name = self.team_table.item(row, 1).text()

        reply = QMessageBox.question(
            self, "Confirm", f"Delete team member '{name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            try:
                delete_resource(resource_id, username=self.user["username"], db_path=self.db_path)
                self._refresh_team_list()
                self._set_status("Team member deleted")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete:\n{e}")

    def _import_team_from_excel(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Select User Database Excel File", "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        if not filepath:
            return

        try:
            import openpyxl
        except ImportError:
            QMessageBox.critical(
                self, "Missing Dependency",
                "openpyxl is required for Excel import.\n\nInstall with:\npip install openpyxl"
            )
            return

        try:
            wb = openpyxl.load_workbook(filepath, read_only=True)
            ws = wb.active

            headers = {}
            header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            for idx, cell in enumerate(header_row):
                if cell:
                    headers[cell.strip().upper()] = idx

            name_col = headers.get("NAME")
            if name_col is None:
                QMessageBox.critical(self, "Invalid Format", "Could not find 'NAME' column.")
                return

            imported = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[name_col]:
                    continue
                name = str(row[name_col]).strip()
                if not name:
                    continue

                try:
                    add_resource(
                        name=name, role="Other", department="",
                        email="", allocation_pct=100.0,
                        username=self.user["username"], db_path=self.db_path
                    )
                    imported += 1
                except Exception:
                    pass

            wb.close()
            self._refresh_team_list()
            QMessageBox.information(self, "Import Complete", f"Imported {imported} team members.")
            self._set_status(f"Imported {imported} team members from Excel")

        except Exception as e:
            QMessageBox.critical(self, "Import Error", f"Failed to import from Excel:\n\n{str(e)}")

    def _build_assignments_subtab(self, tab):
        layout = QVBoxLayout(tab)

        # Project selector
        selector_layout = QHBoxLayout()
        selector_layout.addWidget(QLabel("Project:"))

        self.assign_project_combo = QComboBox()
        self.assign_project_combo.setMinimumWidth(300)
        self.assign_project_combo.currentIndexChanged.connect(self._on_assign_project_change)
        selector_layout.addWidget(self.assign_project_combo)

        if self.can_edit:
            assign_btn = QPushButton("➕ Assign Member")
            assign_btn.clicked.connect(self._assign_member_to_project)
            selector_layout.addWidget(assign_btn)

            remove_btn = QPushButton("🗑️ Remove")
            remove_btn.clicked.connect(self._remove_member_from_project)
            selector_layout.addWidget(remove_btn)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_assignments)
        selector_layout.addWidget(refresh_btn)

        selector_layout.addStretch()
        layout.addLayout(selector_layout)

        # Assignments table
        self.assign_table = QTableWidget()
        self.assign_table.setColumnCount(6)
        self.assign_table.setHorizontalHeaderLabels(["ID", "Name", "Role", "Project Role", "Allocation %", "Notes"])
        self.assign_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.assign_table.setAlternatingRowColors(True)
        self.assign_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.assign_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.assign_table)

        self._assign_projects_cache = []
        self._refresh_assignment_projects()

    def _refresh_assignment_projects(self):
        with db._connect(self.db_path) as conn:
            rows = conn.execute("SELECT id, name FROM projects ORDER BY name").fetchall()
        self._assign_projects_cache = list(rows)
        self.assign_project_combo.clear()
        self.assign_project_combo.addItems([f"{r['id']} — {r['name']}" for r in rows])
        if rows:
            self.assign_project_combo.setCurrentIndex(0)
            self._on_assign_project_change()

    def _on_assign_project_change(self, index=None):
        idx = self.assign_project_combo.currentIndex()
        if idx < 0 or not self._assign_projects_cache:
            return

        project_id = self._assign_projects_cache[idx]["id"]
        self.assign_table.setRowCount(0)

        try:
            assignments = get_project_assignments(project_id, self.db_path)
            for a in assignments:
                row = self.assign_table.rowCount()
                self.assign_table.insertRow(row)
                values = [
                    str(a["resource_id"]), a["name"], a["role"],
                    a["role_in_project"], f"{a['allocation_pct']:.0f}%", a["notes"]
                ]
                for col, val in enumerate(values):
                    self.assign_table.setItem(row, col, QTableWidgetItem(val))
        except Exception as e:
            print(f"Error loading assignments: {e}")

    def _refresh_assignments(self):
        self._refresh_assignment_projects()

    def _assign_member_to_project(self):
        idx = self.assign_project_combo.currentIndex()
        if idx < 0:
            QMessageBox.warning(self, "Select", "Please select a project first.")
            return

        project_id = self._assign_projects_cache[idx]["id"]
        project_name = self._assign_projects_cache[idx]["name"]

        dialog = _AssignmentDialog(self, f"Assign to: {project_name}", self.db_path)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                assign_resource_to_project(
                    project_id=project_id,
                    resource_id=dialog.result["resource_id"],
                    role_in_project=dialog.result.get("role_in_project", ""),
                    allocation_pct=dialog.result.get("allocation_pct", 100),
                    notes=dialog.result.get("notes", ""),
                    username=self.user["username"],
                    db_path=self.db_path
                )
                self._on_assign_project_change()
                self._set_status("Team member assigned to project")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to assign:\n{e}")

    def _remove_member_from_project(self):
        row = self.assign_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Please select an assignment to remove.")
            return

        idx = self.assign_project_combo.currentIndex()
        project_id = self._assign_projects_cache[idx]["id"]
        resource_id = int(self.assign_table.item(row, 0).text())

        reply = QMessageBox.question(
            self, "Confirm", "Remove this assignment?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            try:
                remove_assignment(project_id, resource_id,
                                 username=self.user["username"], db_path=self.db_path)
                self._on_assign_project_change()
                self._set_status("Assignment removed")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to remove:\n{e}")

    def _build_workload_subtab(self, tab):
        layout = QVBoxLayout(tab)

        refresh_btn = QPushButton("🔄 Refresh Workload")
        refresh_btn.clicked.connect(self._refresh_workload)
        layout.addWidget(refresh_btn, alignment=Qt.AlignmentFlag.AlignLeft)

        self.workload_table = QTableWidget()
        self.workload_table.setColumnCount(7)
        self.workload_table.setHorizontalHeaderLabels([
            "Name", "Role", "Capacity %", "Allocated %", "Available %", "Projects", "Status"
        ])
        self.workload_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.workload_table.setAlternatingRowColors(True)
        self.workload_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.workload_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.workload_table)

        self._refresh_workload()

    def _refresh_workload(self):
        self.workload_table.setRowCount(0)
        try:
            summary = get_team_utilization_summary(self.db_path)

            for util in summary:
                if "error" in util:
                    continue

                row = self.workload_table.rowCount()
                self.workload_table.insertRow(row)

                status_display = {
                    "under": "🟡 Under",
                    "optimal": "🟢 Optimal",
                    "over": "🔴 Over",
                }.get(util["status"], util["status"])

                values = [
                    util["resource_name"], "",
                    f"{util['max_capacity']:.0f}%",
                    f"{util['total_allocation']:.0f}%",
                    f"{util['available_capacity']:.0f}%",
                    str(util["project_count"]),
                    status_display,
                ]

                for col, val in enumerate(values):
                    self.workload_table.setItem(row, col, QTableWidgetItem(val))

        except Exception as e:
            print(f"Error loading workload: {e}")

    # ═════════════════════════════════════════════════════════════════════
    # TAB 7 — ADMIN
    # ═════════════════════════════════════════════════════════════════════

    def _build_admin_tab(self):
        layout = QVBoxLayout(self.tab_admin)
        layout.setContentsMargins(10, 10, 10, 10)

        # Header
        header_label = QLabel("Administration")
        header_label.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        layout.addWidget(header_label)

        # Admin sub-notebook
        admin_notebook = QTabWidget()

        # Users subtab
        users_tab = QWidget()
        admin_notebook.addTab(users_tab, "  👤 Users  ")
        self._build_users_subtab(users_tab)

        # Audit subtab
        audit_tab = QWidget()
        admin_notebook.addTab(audit_tab, "  📋 Audit Log  ")
        self._build_audit_subtab(audit_tab)

        # Reference Lines subtab
        ref_tab = QWidget()
        admin_notebook.addTab(ref_tab, "  📏 Reference Lines  ")
        self._build_reflines_subtab(ref_tab)

        # Backup subtab
        backup_tab = QWidget()
        admin_notebook.addTab(backup_tab, "  💾 Backup & Restore  ")
        self._build_backup_subtab(backup_tab)

        layout.addWidget(admin_notebook)

    def _build_users_subtab(self, tab):
        layout = QVBoxLayout(tab)

        # Buttons
        btn_layout = QHBoxLayout()
        add_btn = QPushButton("➕ Add User")
        add_btn.setObjectName("accentBtn")
        add_btn.clicked.connect(self._add_user)
        btn_layout.addWidget(add_btn)

        role_btn = QPushButton("🔄 Change Role")
        role_btn.clicked.connect(self._change_user_role)
        btn_layout.addWidget(role_btn)

        del_btn = QPushButton("🗑️ Delete User")
        del_btn.clicked.connect(self._delete_user)
        btn_layout.addWidget(del_btn)

        btn_layout.addStretch()

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_users)
        btn_layout.addWidget(refresh_btn)

        layout.addLayout(btn_layout)

        # Users table
        self.users_table = QTableWidget()
        self.users_table.setColumnCount(4)
        self.users_table.setHorizontalHeaderLabels(["Username", "Role", "Full Name", "Created"])
        self.users_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.users_table.setAlternatingRowColors(True)
        self.users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.users_table)

        self._refresh_users()

    def _refresh_users(self):
        self.users_table.setRowCount(0)
        for u in auth.list_users(self.db_path):
            row = self.users_table.rowCount()
            self.users_table.insertRow(row)
            values = [u["username"], u["role"], u["full_name"] or "", u["created_at"]]
            for col, val in enumerate(values):
                self.users_table.setItem(row, col, QTableWidgetItem(val))

    def _add_user(self):
        dialog = _UserDialog(self, "Add User")
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                auth.create_user(**dialog.result, db_path=self.db_path)
                self._refresh_users()
                self._set_status(f"User '{dialog.result['username']}' created")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _change_user_role(self):
        row = self.users_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Select a user first.")
            return

        username = self.users_table.item(row, 0).text()
        new_role, ok = QInputDialog.getText(
            self, "Change Role",
            f"New role for '{username}':\n(admin / editor / viewer)"
        )
        if ok and new_role:
            try:
                auth.update_user_role(username, new_role.strip(), self.user["username"], self.db_path)
                self._refresh_users()
                self._set_status(f"Role for '{username}' changed to '{new_role.strip()}'")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _delete_user(self):
        row = self.users_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Select a user first.")
            return

        username = self.users_table.item(row, 0).text()
        if username == self.user["username"]:
            QMessageBox.warning(self, "Error", "Cannot delete yourself.")
            return

        reply = QMessageBox.question(
            self, "Confirm", f"Delete user '{username}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            auth.delete_user(username, self.user["username"], self.db_path)
            self._refresh_users()
            self._set_status(f"User '{username}' deleted")

    def _build_audit_subtab(self, tab):
        layout = QVBoxLayout(tab)

        # Filter bar
        filter_group = QGroupBox("  🔍 Filters  ")
        filter_layout = QGridLayout(filter_group)

        filter_layout.addWidget(QLabel("User:"), 0, 0)
        self.audit_user_combo = QComboBox()
        filter_layout.addWidget(self.audit_user_combo, 0, 1)

        filter_layout.addWidget(QLabel("Action:"), 0, 2)
        self.audit_action_combo = QComboBox()
        filter_layout.addWidget(self.audit_action_combo, 0, 3)

        filter_layout.addWidget(QLabel("From:"), 0, 4)
        self.audit_from_entry = QLineEdit((datetime.date.today() - datetime.timedelta(days=30)).strftime("%Y-%m-%d"))
        self.audit_from_entry.setMaximumWidth(100)
        filter_layout.addWidget(self.audit_from_entry, 0, 5)

        filter_layout.addWidget(QLabel("To:"), 0, 6)
        self.audit_to_entry = QLineEdit(datetime.date.today().strftime("%Y-%m-%d"))
        self.audit_to_entry.setMaximumWidth(100)
        filter_layout.addWidget(self.audit_to_entry, 0, 7)

        filter_layout.addWidget(QLabel("Search:"), 1, 0)
        self.audit_search_entry = QLineEdit()
        self.audit_search_entry.returnPressed.connect(self._refresh_audit)
        filter_layout.addWidget(self.audit_search_entry, 1, 1, 1, 3)

        apply_btn = QPushButton("🔍 Apply Filters")
        apply_btn.clicked.connect(self._refresh_audit)
        filter_layout.addWidget(apply_btn, 1, 4)

        clear_btn = QPushButton("🔄 Clear Filters")
        clear_btn.clicked.connect(self._clear_audit_filters)
        filter_layout.addWidget(clear_btn, 1, 5)

        export_btn = QPushButton("📤 Export CSV")
        export_btn.clicked.connect(self._export_audit_csv)
        filter_layout.addWidget(export_btn, 1, 6)

        layout.addWidget(filter_group)

        # Audit table
        self.audit_table = QTableWidget()
        self.audit_table.setColumnCount(5)
        self.audit_table.setHorizontalHeaderLabels(["Icon", "Timestamp", "User", "Action", "Details"])
        self.audit_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.audit_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.audit_table.setAlternatingRowColors(True)
        self.audit_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.audit_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.audit_table)

        self.audit_status = QLabel("")
        self.audit_status.setStyleSheet("color: #666666;")
        layout.addWidget(self.audit_status)

        self._init_audit_filters()
        self._refresh_audit()

    def _init_audit_filters(self):
        try:
            users = get_unique_users(self.db_path)
            self.audit_user_combo.clear()
            self.audit_user_combo.addItem("All")
            self.audit_user_combo.addItems(users)

            actions = get_unique_actions(self.db_path)
            self.audit_action_combo.clear()
            self.audit_action_combo.addItem("All")
            self.audit_action_combo.addItems(actions)
        except Exception as e:
            print(f"Error initializing audit filters: {e}")

    def _refresh_audit(self):
        self.audit_table.setRowCount(0)

        try:
            username = self.audit_user_combo.currentText()
            if username == "All":
                username = None

            action = self.audit_action_combo.currentText()
            if action == "All":
                action = None

            start_date = None
            end_date = None
            try:
                from_str = self.audit_from_entry.text().strip()
                if from_str:
                    start_date = datetime.datetime.strptime(from_str, "%Y-%m-%d").date()
            except ValueError:
                pass

            try:
                to_str = self.audit_to_entry.text().strip()
                if to_str:
                    end_date = datetime.datetime.strptime(to_str, "%Y-%m-%d").date()
            except ValueError:
                pass

            search_term = self.audit_search_entry.text().strip() or None

            entries = get_audit_log(
                db_path=self.db_path, limit=1000,
                username=username, action=action,
                start_date=start_date, end_date=end_date,
                search_term=search_term,
            )

            for entry in entries:
                row = self.audit_table.rowCount()
                self.audit_table.insertRow(row)

                icon = get_action_icon(entry["action"])
                values = [icon, entry["timestamp"], entry["username"], entry["action"], entry["detail"] or ""]
                for col, val in enumerate(values):
                    self.audit_table.setItem(row, col, QTableWidgetItem(val))

            self.audit_status.setText(f"Showing {len(entries)} entries")

        except Exception as e:
            self.audit_status.setText(f"Error: {e}")

    def _clear_audit_filters(self):
        self.audit_user_combo.setCurrentText("All")
        self.audit_action_combo.setCurrentText("All")
        self.audit_from_entry.setText((datetime.date.today() - datetime.timedelta(days=30)).strftime("%Y-%m-%d"))
        self.audit_to_entry.setText(datetime.date.today().strftime("%Y-%m-%d"))
        self.audit_search_entry.clear()
        self._refresh_audit()

    def _export_audit_csv(self):
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Export Audit Log",
            f"audit_log_{datetime.date.today().strftime('%Y%m%d')}.csv",
            "CSV files (*.csv)"
        )

        if not filepath:
            return

        try:
            username = self.audit_user_combo.currentText()
            if username == "All":
                username = None

            action = self.audit_action_combo.currentText()
            if action == "All":
                action = None

            start_date = None
            end_date = None
            try:
                from_str = self.audit_from_entry.text().strip()
                if from_str:
                    start_date = datetime.datetime.strptime(from_str, "%Y-%m-%d").date()
                to_str = self.audit_to_entry.text().strip()
                if to_str:
                    end_date = datetime.datetime.strptime(to_str, "%Y-%m-%d").date()
            except ValueError:
                pass

            search_term = self.audit_search_entry.text().strip() or None

            count = export_audit_log_csv(
                output_path=pathlib.Path(filepath),
                db_path=self.db_path,
                username=username, action=action,
                start_date=start_date, end_date=end_date,
                search_term=search_term,
            )

            self._set_status(f"Exported {count} entries to CSV")
            QMessageBox.information(self, "Export Complete", f"Exported {count} audit entries to:\n\n{filepath}")

        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export:\n{e}")

    def _build_reflines_subtab(self, tab):
        layout = QVBoxLayout(tab)

        # Buttons
        btn_layout = QHBoxLayout()
        add_btn = QPushButton("➕ Add Reference Line")
        add_btn.setObjectName("accentBtn")
        add_btn.clicked.connect(self._add_ref_line)
        btn_layout.addWidget(add_btn)

        del_btn = QPushButton("🗑️ Delete")
        del_btn.clicked.connect(self._delete_ref_line)
        btn_layout.addWidget(del_btn)

        btn_layout.addStretch()

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_ref_lines)
        btn_layout.addWidget(refresh_btn)

        layout.addLayout(btn_layout)

        # Reference lines table
        self.ref_table = QTableWidget()
        self.ref_table.setColumnCount(5)
        self.ref_table.setHorizontalHeaderLabels(["ID", "Name", "Date", "Color", "Line Style"])
        self.ref_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.ref_table.setAlternatingRowColors(True)
        self.ref_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.ref_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.ref_table)

        self._refresh_ref_lines()

    def _refresh_ref_lines(self):
        self.ref_table.setRowCount(0)
        with db._connect(self.db_path) as conn:
            rows = conn.execute("SELECT * FROM reference_lines ORDER BY date").fetchall()
        for r in rows:
            row = self.ref_table.rowCount()
            self.ref_table.insertRow(row)
            values = [str(r["id"]), r["name"], r["date"], r["color"], r["style"]]
            for col, val in enumerate(values):
                self.ref_table.setItem(row, col, QTableWidgetItem(val))

    def _add_ref_line(self):
        dialog = _RefLineDialog(self, "Add Reference Line")
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            try:
                db.add_reference_line(**dialog.result, username=self.user["username"], db_path=self.db_path)
                self._refresh_ref_lines()
                self._set_status("Reference line added")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _delete_ref_line(self):
        row = self.ref_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Select", "Select a reference line first.")
            return

        ref_id = int(self.ref_table.item(row, 0).text())
        reply = QMessageBox.question(
            self, "Confirm", "Delete this reference line?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            db.delete_reference_line(ref_id, username=self.user["username"], db_path=self.db_path)
            self._refresh_ref_lines()

    def _build_backup_subtab(self, tab):
        layout = QVBoxLayout(tab)

        # Buttons
        btn_layout = QHBoxLayout()
        create_btn = QPushButton("📦 Create Backup")
        create_btn.clicked.connect(self._create_manual_backup)
        btn_layout.addWidget(create_btn)

        restore_btn = QPushButton("📂 Restore Selected")
        restore_btn.clicked.connect(self._restore_selected_backup)
        btn_layout.addWidget(restore_btn)

        delete_btn = QPushButton("🗑️ Delete Selected")
        delete_btn.clicked.connect(self._delete_selected_backup)
        btn_layout.addWidget(delete_btn)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_backups)
        btn_layout.addWidget(refresh_btn)

        btn_layout.addSpacing(20)

        export_json_btn = QPushButton("📤 Export to JSON")
        export_json_btn.clicked.connect(self._export_json)
        btn_layout.addWidget(export_json_btn)

        import_json_btn = QPushButton("📥 Import from JSON")
        import_json_btn.clicked.connect(self._import_json_backup)
        btn_layout.addWidget(import_json_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Backup list
        layout.addWidget(QLabel("Available Backups:"))

        self.backup_table = QTableWidget()
        self.backup_table.setColumnCount(3)
        self.backup_table.setHorizontalHeaderLabels(["Backup Name", "Created", "Size (KB)"])
        self.backup_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.backup_table.setAlternatingRowColors(True)
        self.backup_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.backup_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.backup_table)

        self._refresh_backups()

    def _refresh_backups(self):
        self.backup_table.setRowCount(0)
        try:
            backups = list_backups(self.db_path or db.DEFAULT_DB_PATH)
            for backup in backups:
                row = self.backup_table.rowCount()
                self.backup_table.insertRow(row)
                values = [
                    backup["name"],
                    backup["created"].strftime("%Y-%m-%d %H:%M:%S"),
                    f"{backup['size_kb']:.1f}",
                ]
                for col, val in enumerate(values):
                    self.backup_table.setItem(row, col, QTableWidgetItem(val))
        except Exception as e:
            print(f"Error loading backups: {e}")

    def _create_manual_backup(self):
        try:
            backup_path = create_backup(self.db_path or db.DEFAULT_DB_PATH, backup_name="manual")
            self._set_status(f"Backup created: {backup_path.name}")
            QMessageBox.information(self, "Backup Created", f"Backup saved successfully!\n\n{backup_path}")
            self._refresh_backups()
        except Exception as e:
            QMessageBox.critical(self, "Backup Error", f"Failed to create backup:\n\n{e}")

    def _restore_selected_backup(self):
        row = self.backup_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "No Selection", "Please select a backup to restore.")
            return

        backup_name = self.backup_table.item(row, 0).text()

        reply = QMessageBox.question(
            self, "Confirm Restore",
            f"Are you sure you want to restore from:\n\n{backup_name}\n\n"
            "This will replace the current database. A pre-restore backup will be created.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            backup_dir = get_backup_dir(self.db_path or db.DEFAULT_DB_PATH)
            backup_path = backup_dir / f"{backup_name}.db"
            restore_backup(backup_path, self.db_path or db.DEFAULT_DB_PATH)
            self._set_status("Database restored successfully")
            QMessageBox.information(self, "Restore Complete", "Database restored successfully!\n\nPlease restart the application.")
            self._refresh_backups()
        except Exception as e:
            QMessageBox.critical(self, "Restore Error", f"Failed to restore backup:\n\n{e}")

    def _delete_selected_backup(self):
        row = self.backup_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "No Selection", "Please select a backup to delete.")
            return

        backup_name = self.backup_table.item(row, 0).text()

        reply = QMessageBox.question(
            self, "Confirm Delete", f"Delete backup:\n\n{backup_name}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            backup_dir = get_backup_dir(self.db_path or db.DEFAULT_DB_PATH)
            backup_path = backup_dir / f"{backup_name}.db"
            delete_backup(backup_path)
            self._set_status(f"Backup deleted: {backup_name}")
            self._refresh_backups()
        except Exception as e:
            QMessageBox.critical(self, "Delete Error", f"Failed to delete backup:\n\n{e}")

    def _export_json(self):
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Export to JSON",
            f"project_data_{datetime.date.today().strftime('%Y%m%d')}.json",
            "JSON files (*.json)"
        )

        if not filepath:
            return

        try:
            export_to_json(self.db_path or db.DEFAULT_DB_PATH, pathlib.Path(filepath))
            self._set_status(f"Exported to {filepath}")
            QMessageBox.information(self, "Export Complete", f"Data exported successfully!\n\n{filepath}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export:\n\n{e}")

    def _import_json_backup(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Import from JSON", "",
            "JSON files (*.json)"
        )

        if not filepath:
            return

        reply = QMessageBox.question(
            self, "Import Mode",
            "How do you want to import?\n\n"
            "YES = Merge with existing data\n"
            "NO = Replace all data (backup will be created)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
        )

        if reply == QMessageBox.StandardButton.Cancel:
            return

        merge = (reply == QMessageBox.StandardButton.Yes)

        try:
            if not merge:
                create_backup(self.db_path or db.DEFAULT_DB_PATH, backup_name="pre_import")

            stats = import_from_json(pathlib.Path(filepath), self.db_path or db.DEFAULT_DB_PATH, merge=merge)

            self._set_status("Import complete")
            QMessageBox.information(
                self, "Import Complete",
                f"Data imported successfully!\n\n"
                f"Projects: {stats['projects']}\n"
                f"Milestones: {stats['milestones']}\n"
                f"Phases: {stats['phases']}\n"
                f"Tasks: {stats['tasks']}"
            )

            self._refresh_backups()
            self._refresh_dashboard()

        except Exception as e:
            QMessageBox.critical(self, "Import Error", f"Failed to import:\n\n{e}")


# ─────────────────────────────────────────────────────────────────────────
# Dialog Classes
# ─────────────────────────────────────────────────────────────────────────

class _ProjectDialog(QDialog):
    def __init__(self, parent, title, initial=None):
        super().__init__(parent)
        self.result = None
        self.setWindowTitle(title)
        self.setFixedSize(380, 400)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        # Name
        layout.addWidget(QLabel("Name:"))
        self.name_entry = QLineEdit()
        layout.addWidget(self.name_entry)

        # Start Date
        layout.addWidget(QLabel("Start Date (YYYY-MM-DD):"))
        self.start_entry = QLineEdit()
        layout.addWidget(self.start_entry)

        # End Date
        layout.addWidget(QLabel("End Date (YYYY-MM-DD):"))
        self.end_entry = QLineEdit()
        layout.addWidget(self.end_entry)

        # Status (auto-calculated)
        layout.addWidget(QLabel("Status:"))
        status_info = QLabel("⚡ Auto-calculated from milestone task completion")
        status_info.setStyleSheet("color: #888888; font-style: italic;")
        layout.addWidget(status_info)

        # Dev Region
        layout.addWidget(QLabel("Development Region:"))
        self.dev_region_combo = QComboBox()
        self.dev_region_combo.addItems(REGION_OPTIONS)
        layout.addWidget(self.dev_region_combo)

        # Sales Region
        layout.addWidget(QLabel("Sales Region:"))
        self.sales_region_combo = QComboBox()
        self.sales_region_combo.addItems(REGION_OPTIONS)
        layout.addWidget(self.sales_region_combo)

        layout.addStretch()

        # Save button
        save_btn = QPushButton("💾 Save")
        save_btn.setObjectName("accentBtn")
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

        if initial:
            self.name_entry.setText(initial.get("name", ""))
            self.start_entry.setText(initial.get("start_date", ""))
            self.end_entry.setText(initial.get("end_date", ""))
            self.dev_region_combo.setCurrentText(initial.get("dev_region", ""))
            self.sales_region_combo.setCurrentText(initial.get("sales_region", ""))

    def _save(self):
        name = self.name_entry.text().strip()
        start = self.start_entry.text().strip()
        end = self.end_entry.text().strip()

        if not name or not start or not end:
            QMessageBox.warning(self, "Required", "Name, Start Date, and End Date are required.")
            return

        self.result = {
            "name": name, "start_date": start, "end_date": end,
            "color": cfg.DEFAULT_PROJECT_COLOR,
            "status": "on-track",
            "dev_region": self.dev_region_combo.currentText(),
            "sales_region": self.sales_region_combo.currentText(),
        }
        self.accept()


class _MilestoneDialog(QDialog):
    MILESTONE_NAMES = [
        "IM", "CM", "PM", "SFM", "SHRM", "Post SHRM",
        "Mule Build", "X0", "MPRM", "X1", "X2", "LRM", "X3", "SOP",
        "Other",
    ]

    def __init__(self, parent, title, initial=None):
        super().__init__(parent)
        self.result = None
        self.setWindowTitle(title)
        self.setFixedSize(350, 250)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        # Milestone Name
        layout.addWidget(QLabel("Milestone Name:"))
        self.name_combo = QComboBox()
        self.name_combo.addItems(self.MILESTONE_NAMES)
        self.name_combo.currentTextChanged.connect(self._on_name_change)
        layout.addWidget(self.name_combo)

        # Other name entry (hidden by default)
        self.other_label = QLabel("Custom Milestone Name:")
        self.other_entry = QLineEdit()
        self.other_label.setVisible(False)
        self.other_entry.setVisible(False)
        layout.addWidget(self.other_label)
        layout.addWidget(self.other_entry)

        # Date
        layout.addWidget(QLabel("Date (YYYY-MM-DD):"))
        self.date_entry = QLineEdit()
        layout.addWidget(self.date_entry)

        layout.addStretch()

        # Save button
        save_btn = QPushButton("💾 Save")
        save_btn.setObjectName("accentBtn")
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

        if initial:
            name = initial.get("name", "")
            if name in self.MILESTONE_NAMES:
                self.name_combo.setCurrentText(name)
                if name == "Other":
                    self._show_other_entry()
            else:
                self.name_combo.setCurrentText("Other")
                self._show_other_entry()
                self.other_entry.setText(name)
            self.date_entry.setText(initial.get("date", ""))

    def _on_name_change(self, text):
        if text == "Other":
            self._show_other_entry()
        else:
            self._hide_other_entry()

    def _show_other_entry(self):
        self.other_label.setVisible(True)
        self.other_entry.setVisible(True)

    def _hide_other_entry(self):
        self.other_label.setVisible(False)
        self.other_entry.setVisible(False)
        self.other_entry.clear()

    def _save(self):
        combo_val = self.name_combo.currentText()
        if combo_val == "Other":
            name = self.other_entry.text().strip()
        else:
            name = combo_val
        date = self.date_entry.text().strip()

        if not name or not date:
            QMessageBox.warning(self, "Required", "Both fields are required.")
            return

        self.result = {"name": name, "date": date}
        self.accept()


class _PhaseDialog(QDialog):
    def __init__(self, parent, title, initial=None):
        super().__init__(parent)
        self.result = None
        self.setWindowTitle(title)
        self.setFixedSize(350, 250)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        layout.addWidget(QLabel("Phase Name:"))
        self.name_entry = QLineEdit()
        layout.addWidget(self.name_entry)

        layout.addWidget(QLabel("Start Date (YYYY-MM-DD):"))
        self.start_entry = QLineEdit()
        layout.addWidget(self.start_entry)

        layout.addWidget(QLabel("End Date (YYYY-MM-DD):"))
        self.end_entry = QLineEdit()
        layout.addWidget(self.end_entry)

        layout.addStretch()

        save_btn = QPushButton("💾 Save")
        save_btn.setObjectName("accentBtn")
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

        if initial:
            self.name_entry.setText(initial.get("name", ""))
            self.start_entry.setText(initial.get("start_date", ""))
            self.end_entry.setText(initial.get("end_date", ""))

    def _save(self):
        name = self.name_entry.text().strip()
        start = self.start_entry.text().strip()
        end = self.end_entry.text().strip()

        if not name or not start or not end:
            QMessageBox.warning(self, "Required", "All fields are required.")
            return

        self.result = {"name": name, "start_date": start, "end_date": end}
        self.accept()


class _UserDialog(QDialog):
    def __init__(self, parent, title):
        super().__init__(parent)
        self.result = None
        self.setWindowTitle(title)
        self.setFixedSize(350, 300)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        layout.addWidget(QLabel("Username:"))
        self.user_entry = QLineEdit()
        layout.addWidget(self.user_entry)

        layout.addWidget(QLabel("Full Name:"))
        self.name_entry = QLineEdit()
        layout.addWidget(self.name_entry)

        layout.addWidget(QLabel("Password:"))
        self.pass_entry = QLineEdit()
        self.pass_entry.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.pass_entry)

        layout.addWidget(QLabel("Role:"))
        self.role_combo = QComboBox()
        self.role_combo.addItems(["admin", "editor", "viewer"])
        self.role_combo.setCurrentText("viewer")
        layout.addWidget(self.role_combo)

        layout.addStretch()

        save_btn = QPushButton("👤 Create User")
        save_btn.setObjectName("accentBtn")
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

    def _save(self):
        username = self.user_entry.text().strip()
        password = self.pass_entry.text().strip()

        if not username or not password:
            QMessageBox.warning(self, "Required", "Username and password are required.")
            return

        self.result = {
            "username": username,
            "password": password,
            "role": self.role_combo.currentText(),
            "full_name": self.name_entry.text().strip(),
        }
        self.accept()


class _RefLineDialog(QDialog):
    def __init__(self, parent, title):
        super().__init__(parent)
        self.result = None
        self.setWindowTitle(title)
        self.setFixedSize(350, 280)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        layout.addWidget(QLabel("Name:"))
        self.name_entry = QLineEdit()
        layout.addWidget(self.name_entry)

        layout.addWidget(QLabel("Date (YYYY-MM-DD):"))
        self.date_entry = QLineEdit()
        layout.addWidget(self.date_entry)

        layout.addWidget(QLabel("Color (#RRGGBB):"))
        self.color_entry = QLineEdit("#2196F3")
        layout.addWidget(self.color_entry)

        layout.addWidget(QLabel("Line Style:"))
        self.style_combo = QComboBox()
        self.style_combo.addItems(["--", "-.", ":", "-"])
        self.style_combo.setCurrentText("-.")
        layout.addWidget(self.style_combo)

        layout.addStretch()

        save_btn = QPushButton("💾 Save")
        save_btn.setObjectName("accentBtn")
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

    def _save(self):
        name = self.name_entry.text().strip()
        date = self.date_entry.text().strip()

        if not name or not date:
            QMessageBox.warning(self, "Required", "Name and Date are required.")
            return

        self.result = {
            "name": name,
            "date": date,
            "color": self.color_entry.text().strip() or "#2196F3",
            "style": self.style_combo.currentText(),
        }
        self.accept()


class _ResourceDialog(QDialog):
    ROLE_OPTIONS = ["Developer", "Designer", "PM", "QA", "Analyst", "Architect", "Lead", "Other"]

    def __init__(self, parent, title, initial=None):
        super().__init__(parent)
        self.result = None
        self.setWindowTitle(title)
        self.setFixedSize(400, 350)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        layout.addWidget(QLabel("Name:"))
        self.name_entry = QLineEdit()
        layout.addWidget(self.name_entry)

        layout.addWidget(QLabel("Role:"))
        self.role_combo = QComboBox()
        self.role_combo.addItems(self.ROLE_OPTIONS)
        layout.addWidget(self.role_combo)

        layout.addWidget(QLabel("Department:"))
        self.dept_entry = QLineEdit()
        layout.addWidget(self.dept_entry)

        layout.addWidget(QLabel("Email:"))
        self.email_entry = QLineEdit()
        layout.addWidget(self.email_entry)

        layout.addWidget(QLabel("Capacity % (0-100):"))
        self.alloc_entry = QLineEdit("100")
        layout.addWidget(self.alloc_entry)

        layout.addStretch()

        save_btn = QPushButton("💾 Save")
        save_btn.setObjectName("accentBtn")
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

        if initial:
            self.name_entry.setText(initial.get("name", ""))
            self.role_combo.setCurrentText(initial.get("role", ""))
            self.dept_entry.setText(initial.get("department", ""))
            self.email_entry.setText(initial.get("email", ""))
            self.alloc_entry.setText(str(initial.get("allocation_pct", 100)))

    def _save(self):
        name = self.name_entry.text().strip()
        if not name:
            QMessageBox.warning(self, "Required", "Name is required.")
            return

        try:
            alloc = float(self.alloc_entry.text().strip() or "100")
        except ValueError:
            alloc = 100.0

        self.result = {
            "name": name,
            "role": self.role_combo.currentText(),
            "department": self.dept_entry.text().strip(),
            "email": self.email_entry.text().strip(),
            "allocation_pct": alloc,
        }
        self.accept()


class _AssignmentDialog(QDialog):
    def __init__(self, parent, title, db_path):
        super().__init__(parent)
        self.result = None
        self.db_path = db_path
        self.setWindowTitle(title)
        self.setFixedSize(400, 300)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        layout.addWidget(QLabel("Team Member:"))
        self.member_combo = QComboBox()
        self._resources = get_all_resources(db_path)
        self.member_combo.addItems([f"{r.id} — {r.name} ({r.role})" for r in self._resources])
        layout.addWidget(self.member_combo)

        layout.addWidget(QLabel("Role in Project:"))
        self.role_entry = QLineEdit()
        layout.addWidget(self.role_entry)

        layout.addWidget(QLabel("Allocation % for this project:"))
        self.alloc_entry = QLineEdit("100")
        layout.addWidget(self.alloc_entry)

        layout.addWidget(QLabel("Notes:"))
        self.notes_entry = QLineEdit()
        layout.addWidget(self.notes_entry)

        layout.addStretch()

        save_btn = QPushButton("💾 Assign")
        save_btn.setObjectName("accentBtn")
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

    def _save(self):
        idx = self.member_combo.currentIndex()
        if idx < 0:
            QMessageBox.warning(self, "Select", "Please select a team member.")
            return

        try:
            alloc = float(self.alloc_entry.text().strip() or "100")
        except ValueError:
            alloc = 100.0

        self.result = {
            "resource_id": self._resources[idx].id,
            "role_in_project": self.role_entry.text().strip(),
            "allocation_pct": alloc,
            "notes": self.notes_entry.text().strip(),
        }
        self.accept()


class _ActivityDialog(QDialog):
    STATUS_OPTIONS = ["WIP", "Completed"]

    def __init__(self, parent, title, db_path, week=None, year=None, initial=None):
        super().__init__(parent)
        self.result = None
        self.db_path = db_path
        self.setWindowTitle(title)
        self.setFixedSize(550, 500)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        layout.addWidget(QLabel("Activity Name: *"))
        self.name_entry = QLineEdit()
        layout.addWidget(self.name_entry)

        # Date row
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Start Date:"))
        self.start_entry = QLineEdit()
        date_layout.addWidget(self.start_entry)
        date_layout.addWidget(QLabel("End Date:"))
        self.end_entry = QLineEdit()
        date_layout.addWidget(self.end_entry)
        layout.addLayout(date_layout)

        layout.addWidget(QLabel("Time Taken:"))
        self.time_entry = QLineEdit()
        self.time_entry.setPlaceholderText("e.g., 2 hours, 1 day")
        layout.addWidget(self.time_entry)

        layout.addWidget(QLabel("Members Involved:"))
        self.members_entry = QLineEdit()
        layout.addWidget(self.members_entry)

        layout.addWidget(QLabel("Hard Points / Issues:"))
        self.hard_points_entry = QTextEdit()
        self.hard_points_entry.setMaximumHeight(80)
        layout.addWidget(self.hard_points_entry)

        layout.addWidget(QLabel("Status:"))
        self.status_combo = QComboBox()
        self.status_combo.addItems(self.STATUS_OPTIONS)
        layout.addWidget(self.status_combo)

        # Attachment
        attach_layout = QHBoxLayout()
        attach_layout.addWidget(QLabel("Attachment:"))
        self.attach_entry = QLineEdit()
        self.attach_entry.setReadOnly(True)
        attach_layout.addWidget(self.attach_entry)
        browse_btn = QPushButton("📎 Browse")
        browse_btn.clicked.connect(self._browse_attachment)
        attach_layout.addWidget(browse_btn)
        layout.addLayout(attach_layout)

        layout.addStretch()

        save_btn = QPushButton("💾 Save Activity")
        save_btn.setObjectName("accentBtn")
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

        if initial:
            self.name_entry.setText(initial.get("activity_name", ""))
            self.start_entry.setText(initial.get("start_date", ""))
            self.end_entry.setText(initial.get("end_date", ""))
            self.time_entry.setText(initial.get("time_taken", ""))
            self.members_entry.setText(initial.get("members", ""))
            self.hard_points_entry.setPlainText(initial.get("hard_points", ""))
            self.status_combo.setCurrentText(initial.get("status", "WIP"))
            self.attach_entry.setText(initial.get("attachment_path", ""))

    def _browse_attachment(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Select Attachment", "", "All files (*.*)")
        if filepath:
            self.attach_entry.setText(filepath)

    def _save(self):
        name = self.name_entry.text().strip()
        if not name:
            QMessageBox.warning(self, "Required", "Activity name is required.")
            return

        self.result = {
            "activity_name": name,
            "start_date": self.start_entry.text().strip(),
            "end_date": self.end_entry.text().strip(),
            "time_taken": self.time_entry.text().strip(),
            "members": self.members_entry.text().strip(),
            "hard_points": self.hard_points_entry.toPlainText().strip(),
            "status": self.status_combo.currentText(),
            "attachment_path": self.attach_entry.text().strip(),
        }
        self.accept()


class _MilestoneTaskDialog(QDialog):
    """Dialog to view and manage milestone tasks."""

    def __init__(self, parent, milestone_name: str, milestone_id: int,
                 project_name: str, milestone_date: datetime.date,
                 can_edit: bool, username: str, db_path: pathlib.Path | None = None):
        super().__init__(parent)
        self.milestone_id = milestone_id
        self.can_edit = can_edit
        self.username = username
        self.db_path = db_path

        self.setWindowTitle(f"Tasks: {milestone_name}")
        self.setMinimumSize(900, 600)
        self.setModal(True)

        layout = QVBoxLayout(self)

        # Header
        header_layout = QVBoxLayout()
        title = QLabel(f"📌 {milestone_name}")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        header_layout.addWidget(title)

        header_layout.addWidget(QLabel(f"📁 Project: {project_name}"))
        header_layout.addWidget(QLabel(f"📅 Date: {milestone_date.strftime('%b %d, %Y')}"))
        layout.addLayout(header_layout)

        # Tasks
        self.tasks = db.get_milestone_tasks_with_status(milestone_id, db_path)
        self.task_widgets = []

        if not self.tasks:
            no_tasks = QLabel("No tasks defined for this milestone.")
            no_tasks.setStyleSheet("color: #999999;")
            no_tasks.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(no_tasks)
        else:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            content = QWidget()
            content_layout = QVBoxLayout(content)

            for idx, task in enumerate(self.tasks):
                task_widget = self._create_task_row(task, idx)
                content_layout.addWidget(task_widget)

            scroll.setWidget(content)
            layout.addWidget(scroll)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        if self.can_edit and self.tasks:
            save_btn = QPushButton("💾 Save Changes")
            save_btn.setObjectName("accentBtn")
            save_btn.clicked.connect(self._save_changes)
            btn_layout.addWidget(save_btn)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _create_task_row(self, task: dict, idx: int) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet("QFrame { background-color: white; border: 1px solid #E0E0E0; border-radius: 4px; }")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(10, 8, 10, 8)

        # Task number and name
        layout.addWidget(QLabel(f"{idx + 1}. {task['task_name']}"))
        layout.addStretch()

        # Status combo
        status_combo = QComboBox()
        status_combo.addItems(["Yet to Start", "WIP", "Completed", "Not Applicable"])
        status_combo.setCurrentText(task["status"])
        status_combo.setEnabled(self.can_edit)
        layout.addWidget(status_combo)

        # Store for saving
        self.task_widgets.append({
            "task_id": task["id"],
            "status_combo": status_combo,
        })

        return frame

    def _save_changes(self):
        for widget in self.task_widgets:
            new_status = widget["status_combo"].currentText()
            db.update_task_status(widget["task_id"], new_status, self.username, self.db_path)

        QMessageBox.information(self, "Saved", "Task statuses updated successfully!")
        self.accept()


class _SummaryDashboardDialog(QDialog):
    """Summary dashboard dialog."""

    def __init__(self, parent, projects, today):
        super().__init__(parent)
        self.projects = projects
        self.today = today

        self.setWindowTitle("📊 Portfolio Summary Dashboard")
        self.setMinimumSize(1000, 700)
        self.setModal(True)

        layout = QVBoxLayout(self)

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("📊 Portfolio Summary Dashboard")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        header_layout.addWidget(title)
        header_layout.addStretch()
        date_label = QLabel(f"Generated: {today.strftime('%B %d, %Y')}")
        date_label.setStyleSheet("color: #666666;")
        header_layout.addWidget(date_label)
        layout.addLayout(header_layout)

        # Stats
        total = len(projects)
        on_track = sum(1 for p in projects if p.computed_status(today) == "on-track")
        at_risk = sum(1 for p in projects if p.computed_status(today) == "at-risk")
        overdue = sum(1 for p in projects if p.computed_status(today) == "overdue")

        stats_layout = QHBoxLayout()
        stats_data = [
            ("📁", "Total Projects", str(total), "#0067C0"),
            ("🟢", "On Track", str(on_track), "#16A34A"),
            ("🟡", "At Risk", str(at_risk), "#EAB308"),
            ("🔴", "Overdue", str(overdue), "#DC2626"),
        ]

        for icon, label, value, color in stats_data:
            card = QFrame()
            card.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border: 1px solid #E0E0E0;
                    border-radius: 4px;
                    padding: 10px;
                }
            """)
            card_layout = QHBoxLayout(card)

            icon_label = QLabel(icon)
            icon_label.setFont(QFont("Segoe UI", 18))
            card_layout.addWidget(icon_label)

            text_widget = QWidget()
            text_layout = QVBoxLayout(text_widget)
            text_layout.setContentsMargins(0, 0, 0, 0)
            text_layout.setSpacing(0)

            value_label = QLabel(value)
            value_label.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
            value_label.setStyleSheet(f"color: {color};")
            text_layout.addWidget(value_label)

            name_label = QLabel(label)
            name_label.setStyleSheet("color: #666666;")
            text_layout.addWidget(name_label)

            card_layout.addWidget(text_widget)
            stats_layout.addWidget(card)

        layout.addLayout(stats_layout)
        layout.addStretch()

        # Close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)


# ─────────────────────────────────────────────────────────────────────────
# Launch Function
# ─────────────────────────────────────────────────────────────────────────

def launch(db_path=None):
    """Launch the Project Dashboard application."""
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyleSheet(WIN11_STYLE)

    login = LoginWindow(db_path=db_path)
    if login.exec() == QDialog.DialogCode.Accepted and login.user:
        main_win = MainApp(user=login.user, db_path=db_path)
        main_win.show()
        sys.exit(app.exec())


if __name__ == "__main__":
    launch()
