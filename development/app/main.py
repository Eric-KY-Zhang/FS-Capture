from __future__ import annotations

import os
import sys
from pathlib import Path

# PyInstaller --windowed mode sets sys.stdout/stderr to None. Many third-party
# libraries (tqdm, akshare's progress bars, requests) blindly call .write() on
# them and crash. Redirect to a sink BEFORE any other import that might use them.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")
if sys.stdin is None:
    sys.stdin = open(os.devnull, encoding="utf-8")


from loguru import logger
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import QApplication

from app.core.cache import close_cache
from app.core.pdf_renderer import shutdown_renderer
from app.core.settings import Settings, config_path, load_settings
from app.ui.i18n import LanguageManager
from app.ui.main_view import MainView
from app.ui.main_window import MainWindow
from app.ui.onboarding_dialog import OnboardingDialog
from app.ui.styles import dark_palette, light_palette, load_qss


def resource_path(relative: str) -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
    return str(base / relative)


def _setup_logging(settings: Settings) -> None:
    log_dir = settings.log_path()
    logger.remove()
    # In PyInstaller --windowed mode, sys.stderr is None — guard the sink.
    if sys.stderr is not None:
        try:
            logger.add(sys.stderr, level="INFO")
        except (TypeError, ValueError):
            pass
    logger.add(
        log_dir / "filings_atlas.log",
        rotation="5 MB",
        retention=5,
        encoding="utf-8",
        enqueue=True,
        level="DEBUG",
    )


def main() -> int:
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("Filings Atlas")
    app.setApplicationDisplayName("Filings Atlas / 全球披露图谱")
    app.setOrganizationName("Eric Nutshell")
    app.setWindowIcon(QIcon(resource_path("app/assets/filings_atlas.ico")))

    # Default font tuned for Win11 + Chinese rendering.
    font = QFont("Microsoft YaHei", 10)
    font.setStyleStrategy(QFont.PreferAntialias)
    app.setFont(font)

    first_launch = not config_path().exists()
    settings = load_settings()
    LanguageManager.instance().set_language(settings.ui.language)
    _setup_logging(settings)
    logger.info("Filings Atlas starting up")

    palette = light_palette if settings.ui.theme == "light" else dark_palette
    app.setStyleSheet(load_qss(palette))
    app.aboutToQuit.connect(shutdown_renderer)
    app.aboutToQuit.connect(close_cache)

    window = MainWindow(settings)
    view = MainView(settings, parent=window)
    window.set_body(view)
    window.show()

    if first_launch:
        onboarding = OnboardingDialog(window)
        if (
            onboarding.exec() == onboarding.DialogCode.Accepted
            and onboarding.open_settings_requested
        ):
            view._open_settings()

    return app.exec()


def _show_fatal(message: str) -> None:
    """Last-resort error dialog (used when QApplication isn't even running yet)."""
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox

        _ = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(None, "Filings Atlas / 全球披露图谱 启动失败", message)
    except Exception:
        # Fall back to native Windows MessageBox via ctypes
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(
                0,
                message,
                "Filings Atlas / 全球披露图谱 启动失败",
                0x10,
            )
        except Exception:
            pass


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        import traceback

        tb = traceback.format_exc()
        try:
            logger.exception("fatal startup error")
        except Exception:
            pass
        # Always write a crashlog file the user can find
        try:
            from pathlib import Path

            crashlog = (
                Path(getattr(sys, "_MEIPASS", "")).parent / "crash.log"
                if getattr(sys, "frozen", False)
                else Path("crash.log")
            )
            if getattr(sys, "frozen", False):
                crashlog = Path(sys.executable).parent / "crash.log"
            crashlog.write_text(tb, encoding="utf-8")
        except Exception:
            pass
        _show_fatal(f"{type(exc).__name__}: {exc}\n\n详见 crash.log")
        sys.exit(1)
