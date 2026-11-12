"""Lifecycle of the macOS Office add-in host."""

from __future__ import annotations

import sys
import threading
from pathlib import Path

from PyQt6.QtCore import QObject, Qt, QThread, QTimer, pyqtSignal, pyqtSlot
from PyQt6.QtWidgets import QFileDialog, QWidget

from localization.manager import translate as tr

OFFICE_ADDIN_ENABLED_KEY = "office_addin_enabled"


def office_addin_supported() -> bool:
    return sys.platform == "darwin"


class QtFileChooser(QObject):
    """Shows JSON open/save panels on the GUI thread for server worker threads."""

    _requested = pyqtSignal(str, str, str, object)

    def __init__(self, parent: QWidget | None) -> None:
        super().__init__(parent)
        self._parent = parent
        self._requested.connect(self._show, Qt.ConnectionType.QueuedConnection)

    def open_json(self, title: str) -> Path | None:
        return self._ask("open", title, "")

    def save_json(self, title: str, suggested_name: str) -> Path | None:
        return self._ask("save", title, suggested_name)

    def _ask(self, mode: str, title: str, suggested: str) -> Path | None:
        result: dict[str, Path | None] = {}
        done = threading.Event()
        self._requested.emit(mode, title, suggested, (result, done))
        done.wait()
        return result.get("path")

    @pyqtSlot(str, str, str, object)
    def _show(self, mode: str, title: str, suggested: str, reply) -> None:
        result, done = reply
        try:
            filters = tr("JSON 预设 (*.json)")
            start = str(Path.home() / "Documents" / suggested) if suggested else str(Path.home() / "Documents")
            if mode == "open":
                name, _ = QFileDialog.getOpenFileName(self._parent, title, start, filters)
            else:
                name, _ = QFileDialog.getSaveFileName(self._parent, title, start, filters)
                if name and not name.lower().endswith(".json"):
                    name += ".json"
            result["path"] = Path(name) if name else None
        finally:
            done.set()


class _EnableWorker(QThread):
    completed = pyqtSignal(object, str)

    def __init__(self, files: QtFileChooser, product_version: str) -> None:
        super().__init__()
        self._files = files
        self._version = product_version

    def run(self) -> None:
        from integration.office_addin import host

        try:
            result = host.enable(product_version=self._version, files=self._files)
        except host.OfficeAddinError as error:
            self.completed.emit(None, str(error))
        except Exception:
            self.completed.emit(None, tr("Office 插件启动失败，请查看运行日志。"))
        else:
            self.completed.emit(result.server, "")


class OfficeAddinController(QObject):
    def __init__(self, cfg, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.cfg = cfg
        self._files = QtFileChooser(parent)
        self._server = None
        self._worker: _EnableWorker | None = None
        self._callbacks: list = []
        self._shutting_down = False

    def enabled_preference(self) -> bool:
        try:
            return bool(self.cfg.get(OFFICE_ADDIN_ENABLED_KEY, False))
        except Exception:
            return False

    def is_running(self) -> bool:
        return self._server is not None

    def status_text(self) -> str:
        if self._server is not None:
            return tr("Office 插件已就绪：重新打开 Word 或 PowerPoint 后，在“开始 › 加载项”中选择 LaTeXSnipper。")
        if self.enabled_preference():
            return tr("Office 插件已启用，但当前未运行")
        return tr("Office 插件已关闭")

    def apply_startup_preference(self) -> None:
        if office_addin_supported() and self.enabled_preference():
            self._start()

    def set_enabled_async(self, enabled: bool, callback=None) -> None:
        if not office_addin_supported():
            if callback:
                QTimer.singleShot(0, lambda: callback(False, tr("Office 插件仅支持 macOS。")))
            return
        self.cfg.set(OFFICE_ADDIN_ENABLED_KEY, bool(enabled))
        if enabled:
            self._start(callback)
            return
        self._stop()
        if callback:
            QTimer.singleShot(0, lambda: callback(True, self.status_text()))

    def _start(self, callback=None) -> None:
        if callback:
            self._callbacks.append(callback)
        if self._server is not None:
            self._finish(True, self.status_text())
            return
        if self._worker is not None:
            return
        from runtime.product_version import read_product_version

        worker = _EnableWorker(self._files, read_product_version())
        worker.completed.connect(self._on_enabled, Qt.ConnectionType.QueuedConnection)
        self._worker = worker
        worker.start()

    @pyqtSlot(object, str)
    def _on_enabled(self, server, error: str) -> None:
        worker, self._worker = self._worker, None
        if worker is not None:
            worker.wait()
        if self._shutting_down or not self.enabled_preference():
            # Disabled or quitting while the trust prompt was open.
            from integration.office_addin import host

            host.disable(server)
            self._finish(False, tr("Office 插件已关闭"))
            return
        if server is None:
            self.cfg.set(OFFICE_ADDIN_ENABLED_KEY, False)
            self._finish(False, error)
            return
        self._server = server
        self._finish(True, self.status_text())

    def _finish(self, ok: bool, message: str) -> None:
        callbacks, self._callbacks = self._callbacks, []
        for callback in callbacks:
            callback(ok, message)

    def _stop(self) -> None:
        from integration.office_addin import host

        server, self._server = self._server, None
        try:
            host.disable(server)
        except OSError:
            pass

    def shutdown(self) -> None:
        """Stop serving on exit but keep the manifest: it is reused on next launch."""
        self._shutting_down = True
        server, self._server = self._server, None
        if server is not None:
            server.stop()
        if self._worker is not None:
            self._worker.wait(2000)
