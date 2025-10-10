"""System tray launcher for the Report PDF Converter application.

This module boots the Flask application defined in :mod:`app` in a background
thread and provides a small Qt based system tray UI to control the web server.
The file doubles as the PyInstaller entry point when bundling the project as a
Windows executable.
"""

from __future__ import annotations

import socket
import sys
import threading
import webbrowser
from contextlib import closing
from dataclasses import dataclass
from typing import Callable

from werkzeug.serving import BaseWSGIServer, make_server

from app import get_app

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QCursor, QIcon, QPainter, QPen, QPixmap
from PyQt6.QtWidgets import QAction, QApplication, QMenu, QSystemTrayIcon


# Default host and port used by the embedded Flask server.
HOST = "127.0.0.1"
PORT = 80


class ServerError(RuntimeError):
    """Raised when the embedded server encounters an unrecoverable error."""


@dataclass(slots=True)
class ServerState:
    """Runtime configuration for the embedded Flask server."""

    host: str = HOST
    port: int = PORT


class ServerController:
    """Start and stop the Flask application in a background thread."""

    def __init__(self, state: ServerState | None = None) -> None:
        self.state = state or ServerState()
        self._server: BaseWSGIServer | None = None
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._app = get_app()

    # ---- public API -------------------------------------------------
    def start(self) -> None:
        """Start the Flask development server if it is not already running."""

        with self._lock:
            if self._server is not None:
                return
            host, port = self.state.host, self.state.port
            if not _is_port_available(host, port):
                port = _find_free_port(host)
                self.state.port = port

            try:
                self._server = make_server(host, port, self._app)
            except OSError as exc:  # pragma: no cover - defensive branch
                raise ServerError("Failed to initialise the embedded server") from exc

            self._thread = threading.Thread(
                target=self._server.serve_forever,
                daemon=True,
            )
            self._thread.start()

    def stop(self) -> None:
        """Stop the running server, if any."""

        with self._lock:
            if self._server is None:
                return
            self._server.shutdown()
            self._server.server_close()
            self._server = None
            self._thread = None

    def is_running(self) -> bool:
        """Return ``True`` when the server thread is active."""

        with self._lock:
            return self._server is not None


def _is_port_available(host: str, port: int) -> bool:
    with closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
        sock.settimeout(0.2)
        return sock.connect_ex((host, port)) != 0


def _find_free_port(host: str) -> int:
    with closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
        sock.bind((host, 0))
        return sock.getsockname()[1]


def _make_status_icon(active: bool) -> QIcon:
    """Generate a circular status icon indicating the server state."""

    size = 128
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

    color = QColor(0, 160, 0) if active else QColor(160, 0, 0)
    painter.setBrush(color)

    pen = QPen(QColor(255, 255, 255, 230))
    pen.setWidth(6)
    painter.setPen(pen)

    margin = 16
    painter.drawEllipse(margin, margin, size - 2 * margin, size - 2 * margin)
    painter.end()
    return QIcon(pixmap)


def _build_menu(*actions: QAction) -> QMenu:
    if len(actions) != 4:
        raise ValueError("Expected start, stop, open and exit actions")

    menu = QMenu()
    menu.setStyleSheet(
        """
        QMenu {
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #cccccc;
            padding: 2px 0;
        }
        QMenu::icon {
            width: 0px;
        }
        QMenu::item {
            padding: 4px 10px;
            margin: 0;
        }
        QMenu::item:selected {
            background-color: #e6f0ff;
            color: #000000;
        }
        QMenu::separator {
            height: 6px;
            margin: 3px 0;
            background: #e0e0e0;
        }
        """
    )

    start, stop, open_action, exit_action = actions
    for action in actions:
        action.setIconVisibleInMenu(False)

    menu.addAction(start)
    menu.addAction(stop)
    menu.addSeparator()
    menu.addAction(open_action)
    menu.addSeparator()
    menu.addAction(exit_action)
    return menu


def _connect_action(action: QAction, callback: Callable[[], None]) -> None:
    action.triggered.connect(callback)  # type: ignore[arg-type]


def main() -> None:
    qt_app = QApplication(sys.argv)
    if not QSystemTrayIcon.isSystemTrayAvailable():
        print("System tray not available.")
        sys.exit(1)

    server = ServerController()

    tray = QSystemTrayIcon()
    tray.setToolTip("ロボ研報告書作成サーバ (OFF)")
    tray.setIcon(_make_status_icon(active=False))

    act_start = QAction("サーバー開始")
    act_stop = QAction("サーバー停止")
    act_open = QAction("画面を開く")
    act_exit = QAction("終了")

    def start_server() -> None:
        server.start()
        tray.setIcon(_make_status_icon(active=True))
        tray.setToolTip(
            f"ロボ研報告書作成サーバ (ON) http://{server.state.host}:{server.state.port}"
        )

    def stop_server() -> None:
        server.stop()
        tray.setIcon(_make_status_icon(active=False))
        tray.setToolTip("ロボ研報告書作成サーバ (OFF)")

    def open_browser() -> None:
        if not server.is_running():
            start_server()
        webbrowser.open(f"http://{server.state.host}:{server.state.port}/")

    def exit_application() -> None:
        stop_server()
        tray.hide()
        QApplication.quit()

    _connect_action(act_start, start_server)
    _connect_action(act_stop, stop_server)
    _connect_action(act_open, open_browser)
    _connect_action(act_exit, exit_application)

    menu = _build_menu(act_start, act_stop, act_open, act_exit)
    tray.setContextMenu(menu)

    def on_activation(reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason in {
            QSystemTrayIcon.ActivationReason.Trigger,
            QSystemTrayIcon.ActivationReason.DoubleClick,
        }:
            menu.popup(QCursor.pos())

    tray.activated.connect(on_activation)  # type: ignore[arg-type]

    tray.show()
    sys.exit(qt_app.exec())


if __name__ == "__main__":
    main()
