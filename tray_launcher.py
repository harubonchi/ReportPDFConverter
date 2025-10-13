# tray_launcher_qt.py
from __future__ import annotations
import sys
import threading
import webbrowser
import socket
from contextlib import closing

from werkzeug.serving import make_server

# ★ ←ここをあなたのメインファイル名に合わせて変更（拡張子なし）
# 例: メインが main.py なら: from main import get_app
from app import get_app

from PyQt6.QtCore import Qt, QPoint
from PyQt6.QtGui import (
    QIcon, QPainter, QBrush, QPen, QColor, QPixmap, QAction, QCursor
)
from PyQt6.QtWidgets import QApplication, QSystemTrayIcon, QMenu

HOST = "0.0.0.0"
PORT = 80  # 既存UIの想定ポート。競合なら空きポートへ回避


# ---- Flaskサーバ制御 ----
class ServerCtl:
    def __init__(self):
        self._server = None
        self._thread = None
        self._lock = threading.Lock()
        self.host = HOST
        self.port = PORT
        self.app = get_app()

    def is_running(self) -> bool:
        with self._lock:
            return self._server is not None

    def start(self):
        with self._lock:
            if self._server is not None:
                return
            if not self._port_free(self.host, self.port):
                self.port = self._find_free_port(self.host)

            self._server = make_server(self.host, self.port, self.app)
            self._thread = threading.Thread(
                target=self._server.serve_forever, daemon=True
            )
            self._thread.start()

    def stop(self):
        with self._lock:
            if self._server is None:
                return
            self._server.shutdown()
            self._server.server_close()
            self._server = None
            self._thread = None

    @staticmethod
    def _port_free(host, port) -> bool:
        with closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as s:
            s.settimeout(0.2)
            return s.connect_ex((host, port)) != 0

    @staticmethod
    def _find_free_port(host) -> int:
        with closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as s:
            s.bind((host, 0))
            return s.getsockname()[1]


# ---- アイコン（緑=ON / 赤=OFF）をQtで描画 ----
def make_dot_icon(on: bool) -> QIcon:
    size = 128
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

    color = QColor(0, 160, 0) if on else QColor(160, 0, 0)
    brush = QBrush(color)
    pen = QPen(QColor(255, 255, 255, 230))
    pen.setWidth(6)

    painter.setBrush(brush)
    painter.setPen(pen)
    margin = 16
    painter.drawEllipse(margin, margin, size - 2 * margin, size - 2 * margin)
    painter.end()
    return QIcon(pm)


def main():
    app = QApplication(sys.argv)
    QApplication.setQuitOnLastWindowClosed(False)

    if not QSystemTrayIcon.isSystemTrayAvailable():
        print("System tray not available.")
        sys.exit(1)

    ctl = ServerCtl()

    tray = QSystemTrayIcon()
    tray.setToolTip("ロボ研報告書作成サーバ (OFF)")
    tray.setIcon(make_dot_icon(on=False))

    # --- メニュー（日本語） ---
    menu = QMenu()

    # 白背景・ホバー時の選択色・左余白を詰める
    menu.setStyleSheet("""
        QMenu {
            background-color: #ffffff;               /* 白背景 */
            color: #000000;                          /* 通常テキスト色 */
            border: 1px solid #cccccc;               /* 薄い枠線 */
            padding: 2px 0;
        }
        QMenu::icon {
            width: 0px;                              /* 左アイコン列を消す */
        }
        QMenu::item {
            padding: 4px 10px;                       /* 各項目の内側余白（左を詰める） */
            margin: 0;
        }
        QMenu::item:selected {
            background-color: #e6f0ff;               /* ホバー時の青みハイライト */
            color: #000000;
        }
        QMenu::separator {
            height: 6px;
            margin: 3px 0;
            background: #e0e0e0;
        }
    """)

    act_start = QAction("サーバー開始", menu)
    act_stop  = QAction("サーバー停止", menu)
    act_open  = QAction("画面を開く", menu)
    act_exit  = QAction("終了", menu)

    # アイコン非表示
    for a in (act_start, act_stop, act_open, act_exit):
        a.setIconVisibleInMenu(False)

    menu.addAction(act_start)
    menu.addAction(act_stop)
    menu.addSeparator()
    menu.addAction(act_open)
    menu.addSeparator()
    menu.addAction(act_exit)

    tray.setContextMenu(menu)

    # --- アクション動作 ---
    def do_start():
        ctl.start()
        tray.setIcon(make_dot_icon(on=True))
        tray.setToolTip(f"ロボ研報告書作成サーバ (ON) http://{ctl.host}:{ctl.port}")

    def do_stop():
        ctl.stop()
        tray.setIcon(make_dot_icon(on=False))
        tray.setToolTip("ロボ研報告書作成サーバ (OFF)")

    def do_open():
        if not ctl.is_running():
            do_start()
        webbrowser.open(f"http://{ctl.host}:{ctl.port}/")

    def do_exit():
        do_stop()
        tray.hide()
        QApplication.quit()

    act_start.triggered.connect(do_start)
    act_stop.triggered.connect(do_stop)
    act_open.triggered.connect(do_open)
    act_exit.triggered.connect(do_exit)

    # --- 左クリックでメニューを開く ---
    # Windowsでは Trigger（シングル左クリック）/ DoubleClick で拾えます。
    def on_activated(reason: QSystemTrayIcon.ActivationReason):
        if reason in (
            QSystemTrayIcon.ActivationReason.Trigger,       # 左クリック
            QSystemTrayIcon.ActivationReason.DoubleClick,   # 左ダブルクリック
        ):
            # カーソル位置にメニューを表示
            pos = QCursor.pos()
            menu.popup(pos)

    tray.activated.connect(on_activated)

    # 起動と同時にサーバを立ち上げる
    do_start()

    tray.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
