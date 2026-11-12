from __future__ import annotations

from unittest.mock import MagicMock

import pytest

pytest.importorskip("PyQt6.QtWidgets")

from PyQt6.QtWidgets import QApplication

import ui.settings_layout_builder as builder
from ui.settings_window import SettingsWindow

ENGLISH = {
    "运行日志: 关": "Runtime Log: Off",
    "运行日志: 开": "Runtime Log: On",
    "自动化接口: 关": "Automation API: Off",
    "自动化接口: 开": "Automation API: On",
    "Office 插件: 关": "Office Add-in: Off",
    "Office 插件: 开": "Office Add-in: On",
}


class Config(dict):
    def get(self, key, default=None):
        return super().get(key, default)

    def set(self, key, value):
        self[key] = value


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication(["test"])


def open_settings(monkeypatch, *, office_addin: bool, english: bool):
    monkeypatch.setattr(builder, "office_addin_supported", lambda: office_addin)
    if english:
        monkeypatch.setattr(builder, "tr", lambda text, **_: ENGLISH.get(text, text))
    context = MagicMock()
    context.cfg = Config()
    context.office_addin.enabled_preference.return_value = False
    window = SettingsWindow(None, context=context)
    window.show()
    QApplication.processEvents()
    return window


def startup_buttons(window):
    buttons = [
        window.runtime_log_button,
        window.automation_api_button,
        window.office_addin_button,
    ]
    return [button for button in buttons if button is not None]


def assert_fully_visible(window):
    viewport = window.scroll_area.viewport()
    for button in startup_buttons(window):
        right = button.mapTo(viewport, button.rect().topRight()).x()
        assert right <= viewport.width(), f"{button.text()!r} is clipped"
    assert window.content_widget.minimumSizeHint().width() <= viewport.width()


@pytest.mark.parametrize("english", [False, True])
def test_startup_buttons_fit_in_both_states(app, monkeypatch, english):
    window = open_settings(monkeypatch, office_addin=True, english=english)
    try:
        assert len(startup_buttons(window)) == 3
        assert_fully_visible(window)
        for button in startup_buttons(window):
            button.setChecked(True)
        window._sync_startup_action_buttons()
        QApplication.processEvents()
        assert_fully_visible(window)
    finally:
        window.deleteLater()


def test_default_width_is_unchanged_when_content_fits(app, monkeypatch):
    window = open_settings(monkeypatch, office_addin=False, english=False)
    try:
        assert window.width() == builder.SETTINGS_MIN_WIDTH
        assert window.minimumWidth() == builder.SETTINGS_MIN_WIDTH
        assert_fully_visible(window)
    finally:
        window.deleteLater()
