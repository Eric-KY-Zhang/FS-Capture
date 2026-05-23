from __future__ import annotations

from PySide6.QtCore import QPoint, QRect, QSize, Qt, Signal
from PySide6.QtGui import QCursor, QMouseEvent
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.core.settings import Settings
from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager


class _TitleBar(QWidget):
    """Custom draggable title bar with min/max/close buttons."""

    minimize_requested = Signal()
    maximize_requested = Signal()
    close_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("TitleBar")
        self.setFixedHeight(48)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 8, 0)
        layout.setSpacing(0)

        self.logo = QLabel("Filings Atlas")
        self.logo.setObjectName("TitleBarLogo")

        self.subtitle = QLabel(ui_strings.MW_SUBTITLE)
        self.subtitle.setObjectName("TitleBarSubtitle")

        layout.addWidget(self.logo)
        layout.addWidget(self.subtitle)
        layout.addStretch(1)

        self.btn_min = self._make_btn("—", "minimize")
        self.btn_max = self._make_btn("☐", "maximize")
        self.btn_close = self._make_btn("✕", "close")

        self.btn_min.clicked.connect(self.minimize_requested.emit)
        self.btn_max.clicked.connect(self.maximize_requested.emit)
        self.btn_close.clicked.connect(self.close_requested.emit)

        layout.addWidget(self.btn_min)
        layout.addWidget(self.btn_max)
        layout.addWidget(self.btn_close)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def _retranslate(self, _lang: str = "") -> None:
        self.subtitle.setText(ui_strings.MW_SUBTITLE)

    @staticmethod
    def _make_btn(text: str, role: str) -> QPushButton:
        b = QPushButton(text)
        b.setObjectName("WindowControl")
        b.setProperty("role", role)
        b.setCursor(Qt.PointingHandCursor)
        b.setFlat(True)
        return b


class MainWindow(QMainWindow):
    """Frameless rounded main window.

    Outer #WindowRoot widget holds the shadow + rounded border. Inside:
      - title bar (custom, draggable)
      - body widget (set via set_body)

    Window can be dragged from the title bar and resized via 6px edge margins.
    """

    EDGE_MARGIN = 6
    MIN_SIZE = QSize(960, 680)

    def __init__(self, settings: Settings) -> None:
        super().__init__()
        self.settings = settings

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.setMouseTracking(True)
        self.resize(settings.ui.window_width, settings.ui.window_height)
        self.setMinimumSize(self.MIN_SIZE)
        self.setWindowTitle(ui_strings.MW_WINDOW_TITLE)

        # Full-window root. Avoid translucent outer margins; they look like a
        # strange transparent border on Windows.
        outer = QWidget(self)
        outer.setMouseTracking(True)
        outer_layout = QVBoxLayout(outer)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        self._root = QWidget(outer)
        self._root.setObjectName("WindowRoot")
        self._root.setMouseTracking(True)

        outer_layout.addWidget(self._root)

        root_layout = QVBoxLayout(self._root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self._title_bar = _TitleBar(self._root)
        self._title_bar.minimize_requested.connect(self.showMinimized)
        self._title_bar.maximize_requested.connect(self._toggle_maximize)
        self._title_bar.close_requested.connect(self.close)
        root_layout.addWidget(self._title_bar)

        self._body_host = QWidget(self._root)
        self._body_host.setMouseTracking(True)
        self._body_layout = QVBoxLayout(self._body_host)
        self._body_layout.setContentsMargins(0, 0, 0, 0)
        self._body_layout.setSpacing(0)
        root_layout.addWidget(self._body_host, 1)

        self.setCentralWidget(outer)

        # Drag / resize state
        self._drag_offset: QPoint | None = None
        self._resize_edge: str | None = None
        self._resize_start_geom: QRect | None = None
        self._resize_start_pos: QPoint | None = None
        LanguageManager.instance().language_changed.connect(self._retranslate)

    # ---- public API ------------------------------------------------------

    def set_body(self, widget: QWidget) -> None:
        for i in reversed(range(self._body_layout.count())):
            item = self._body_layout.takeAt(i)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        widget.setParent(self._body_host)
        self._body_layout.addWidget(widget)

    def _retranslate(self, _lang: str = "") -> None:
        self.setWindowTitle(ui_strings.MW_WINDOW_TITLE)

    # ---- maximize toggle -------------------------------------------------

    def _toggle_maximize(self) -> None:
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def changeEvent(self, event) -> None:  # type: ignore[override]
        # Keep the frameless window flush with the screen edge when maximized.
        super().changeEvent(event)
        if hasattr(self, "_root"):
            self._root.setStyleSheet("#WindowRoot { border-radius: 0; }")

    # ---- drag + resize ---------------------------------------------------

    def mousePressEvent(self, event: QMouseEvent) -> None:  # type: ignore[override]
        if event.button() == Qt.LeftButton:
            local = event.position().toPoint()
            edge = self._edge_at(local)
            if edge and not self.isMaximized():
                self._resize_edge = edge
                self._resize_start_geom = QRect(self.geometry())
                self._resize_start_pos = event.globalPosition().toPoint()
                event.accept()
                return
            # Title-bar drag area (allow grabbing on the title bar but not on its buttons)
            tb_rect = self._title_bar.geometry()
            tb_local = self._title_bar.mapFromParent(local)
            if tb_rect.contains(local):
                child = self._title_bar.childAt(tb_local)
                if not isinstance(child, QPushButton):
                    if self.isMaximized():
                        return
                    self._drag_offset = (
                        event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                    )
                    event.accept()
                    return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:  # type: ignore[override]
        if (
            self._resize_edge
            and self._resize_start_geom is not None
            and self._resize_start_pos is not None
        ):
            self._perform_resize(event.globalPosition().toPoint())
            event.accept()
            return
        if self._drag_offset is not None and (event.buttons() & Qt.LeftButton):
            self.move(event.globalPosition().toPoint() - self._drag_offset)
            event.accept()
            return
        if not self.isMaximized():
            self._update_cursor(event.position().toPoint())
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:  # type: ignore[override]
        self._drag_offset = None
        self._resize_edge = None
        self._resize_start_geom = None
        self._resize_start_pos = None
        self.unsetCursor()
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:  # type: ignore[override]
        local = event.position().toPoint()
        if self._title_bar.geometry().contains(local):
            tb_local = self._title_bar.mapFromParent(local)
            child = self._title_bar.childAt(tb_local)
            if not isinstance(child, QPushButton):
                self._toggle_maximize()
                event.accept()
                return
        super().mouseDoubleClickEvent(event)

    # ---- helpers ---------------------------------------------------------

    def _edge_at(self, pos: QPoint) -> str | None:
        m = self.EDGE_MARGIN
        x, y = pos.x(), pos.y()
        w, h = self.width(), self.height()
        left, right = x < m, x > w - m
        top, bottom = y < m, y > h - m
        if top and left:
            return "top_left"
        if top and right:
            return "top_right"
        if bottom and left:
            return "bottom_left"
        if bottom and right:
            return "bottom_right"
        if top:
            return "top"
        if bottom:
            return "bottom"
        if left:
            return "left"
        if right:
            return "right"
        return None

    def _update_cursor(self, pos: QPoint) -> None:
        edge = self._edge_at(pos)
        mapping = {
            "top": Qt.SizeVerCursor,
            "bottom": Qt.SizeVerCursor,
            "left": Qt.SizeHorCursor,
            "right": Qt.SizeHorCursor,
            "top_left": Qt.SizeFDiagCursor,
            "bottom_right": Qt.SizeFDiagCursor,
            "top_right": Qt.SizeBDiagCursor,
            "bottom_left": Qt.SizeBDiagCursor,
        }
        if edge:
            self.setCursor(QCursor(mapping[edge]))
        else:
            self.unsetCursor()

    def _perform_resize(self, global_pos: QPoint) -> None:
        if (
            self._resize_edge is None
            or self._resize_start_geom is None
            or self._resize_start_pos is None
        ):
            return
        delta = global_pos - self._resize_start_pos
        g = QRect(self._resize_start_geom)
        edge = self._resize_edge
        if "left" in edge:
            g.setLeft(g.left() + delta.x())
        if "right" in edge:
            g.setRight(g.right() + delta.x())
        if "top" in edge:
            g.setTop(g.top() + delta.y())
        if "bottom" in edge:
            g.setBottom(g.bottom() + delta.y())
        if g.width() < self.MIN_SIZE.width():
            if "left" in edge:
                g.setLeft(g.right() - self.MIN_SIZE.width())
            else:
                g.setRight(g.left() + self.MIN_SIZE.width())
        if g.height() < self.MIN_SIZE.height():
            if "top" in edge:
                g.setTop(g.bottom() - self.MIN_SIZE.height())
            else:
                g.setBottom(g.top() + self.MIN_SIZE.height())
        self.setGeometry(g)
