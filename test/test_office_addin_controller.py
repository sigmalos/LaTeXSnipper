from __future__ import annotations

from types import SimpleNamespace

import pytest

pytest.importorskip("PyQt6.QtWidgets")

from PyQt6.QtCore import QCoreApplication, QEventLoop, QTimer
from PyQt6.QtWidgets import QApplication

from integration.office_addin import host
from ui.controllers import office_addin
from ui.controllers.office_addin import OFFICE_ADDIN_ENABLED_KEY, OfficeAddinController


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication(["test"])


class Config(dict):
    def set(self, key, value):
        self[key] = value


class FakeServer:
    def __init__(self):
        self.stopped = False

    def stop(self):
        self.stopped = True


@pytest.fixture
def fakes(monkeypatch):
    calls = {"enable": 0, "disable": []}
    state = {"error": None}

    def enable(**kwargs):
        calls["enable"] += 1
        if state["error"]:
            raise host.OfficeAddinError(state["error"])
        return SimpleNamespace(server=FakeServer(), manifests=[], trusted_now=False)

    monkeypatch.setattr(host, "enable", enable)
    monkeypatch.setattr(host, "disable", lambda server, **_: calls["disable"].append(server))
    monkeypatch.setattr(office_addin, "office_addin_supported", lambda: True)
    monkeypatch.setattr("runtime.product_version.read_product_version", lambda: "3.3.0")
    return calls, state


def wait_for(callback_results, count=1, timeout=5000):
    loop = QEventLoop()
    timer = QTimer()
    timer.timeout.connect(lambda: len(callback_results) >= count and loop.quit())
    timer.start(10)
    QTimer.singleShot(timeout, loop.quit)
    loop.exec()
    timer.stop()
    QCoreApplication.processEvents()


def test_enable_starts_server_and_remembers_preference(app, fakes):
    calls, _ = fakes
    cfg = Config()
    controller = OfficeAddinController(cfg)
    results = []
    controller.set_enabled_async(True, lambda ok, message: results.append((ok, message)))
    wait_for(results)
    assert results[0][0] is True
    assert controller.is_running()
    assert cfg[OFFICE_ADDIN_ENABLED_KEY] is True
    assert calls["enable"] == 1


def test_failed_enable_reports_error_and_clears_preference(app, fakes):
    _, state = fakes
    state["error"] = "端口被占用"
    cfg = Config()
    controller = OfficeAddinController(cfg)
    results = []
    controller.set_enabled_async(True, lambda ok, message: results.append((ok, message)))
    wait_for(results)
    assert results == [(False, "端口被占用")]
    assert not controller.is_running()
    assert cfg[OFFICE_ADDIN_ENABLED_KEY] is False


def test_disable_stops_server_and_removes_manifest(app, fakes):
    calls, _ = fakes
    cfg = Config()
    controller = OfficeAddinController(cfg)
    results = []
    controller.set_enabled_async(True, lambda ok, message: results.append(ok))
    wait_for(results)
    server = controller._server
    controller.set_enabled_async(False, lambda ok, message: results.append(ok))
    wait_for(results, 2)
    assert results == [True, True]
    assert calls["disable"] == [server]
    assert not controller.is_running()
    assert cfg[OFFICE_ADDIN_ENABLED_KEY] is False


def test_disabling_while_starting_does_not_leave_a_server(app, fakes):
    calls, _ = fakes
    cfg = Config()
    controller = OfficeAddinController(cfg)
    results = []
    controller.set_enabled_async(True, lambda ok, message: results.append(("enable", ok)))
    controller.set_enabled_async(False, lambda ok, message: results.append(("disable", ok)))
    wait_for(results, 2)
    assert ("enable", False) in results
    assert not controller.is_running()
    assert len(calls["disable"]) == 2


def test_shutdown_keeps_manifest_for_next_launch(app, fakes):
    calls, _ = fakes
    controller = OfficeAddinController(Config())
    results = []
    controller.set_enabled_async(True, lambda ok, message: results.append(ok))
    wait_for(results)
    server = controller._server
    controller.shutdown()
    assert server.stopped
    assert calls["disable"] == []


def test_startup_preference_only_applies_on_macos(app, fakes, monkeypatch):
    calls, _ = fakes
    cfg = Config({OFFICE_ADDIN_ENABLED_KEY: True})
    monkeypatch.setattr(office_addin, "office_addin_supported", lambda: False)
    OfficeAddinController(cfg).apply_startup_preference()
    QCoreApplication.processEvents()
    assert calls["enable"] == 0
