"""Enable and disable the macOS Office add-in as one unit.

Enabling creates (or renews) the loopback certificate, asks macOS to trust it,
starts the HTTPS host and sideloads the manifest. Disabling stops the host and
removes only this add-in's manifest; the certificate stays trusted so the next
enable does not prompt again. ``remove_trust`` undoes that explicitly.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from integration.office_addin import certificate, sideload
from integration.office_addin.assets import OfficeAddinAssets, resolve_assets
from integration.office_addin.outlines import SystemFontOutlines
from integration.office_addin.recognition import AutomationRelay
from integration.office_addin.server import DEFAULT_PORT, OfficeAddinServer
from integration.office_addin.services import (
    FileChooser,
    OfficeAddinServiceImpl,
    mathjax_version,
    read_symbol_fonts,
)
from integration.office_addin.typography import SettingsStore

STATE_DIRECTORY = "office-addin"


@dataclass(frozen=True)
class EnableResult:
    server: OfficeAddinServer
    manifests: list[Path]
    trusted_now: bool


class OfficeAddinError(RuntimeError):
    pass


def state_directory() -> Path:
    from runtime.app_paths import app_state_dir

    return app_state_dir() / STATE_DIRECTORY


def enable(
    *,
    product_version: str,
    files: FileChooser | None,
    port: int = DEFAULT_PORT,
    state: Path | None = None,
    assets: OfficeAddinAssets | None = None,
    home: Path | None = None,
) -> EnableResult:
    state = state or state_directory()
    assets = assets or resolve_assets()
    try:
        assets.validate()
    except FileNotFoundError as error:
        raise OfficeAddinError(str(error)) from None
    try:
        identity, created = certificate.ensure_certificate(state)
        trusted_now = False
        if created or not certificate.is_trusted(identity):
            certificate.trust(identity)
            trusted_now = True
        if not certificate.is_trusted(identity):
            raise OfficeAddinError("macOS 未信任 Office 加载项证书，Office 将无法加载插件页面。")
    except certificate.CertificateError as error:
        raise OfficeAddinError(f"无法设置 Office 加载项证书：{error}") from None

    symbol_fonts = read_symbol_fonts(assets.mathjax)
    services = OfficeAddinServiceImpl(
        settings=SettingsStore(state / "settings.json", symbol_fonts),
        outlines=SystemFontOutlines(),
        relay=AutomationRelay(),
        files=files,
        mathjax_version=mathjax_version(assets.mathjax),
    )
    server = OfficeAddinServer(assets=assets, certificate=identity, services=services, port=port)
    try:
        server.start()
    except OSError as error:
        raise OfficeAddinError(f"无法监听 127.0.0.1:{port}，端口可能已被占用。") from error
    try:
        manifest = sideload.render_manifest(
            assets.manifest_template.read_text(encoding="utf-8"),
            base_url=server.base_url,
            product_version=product_version,
        )
        manifests = sideload.install_manifest(manifest, home=home)
    except (OSError, ValueError) as error:
        server.stop()
        raise OfficeAddinError(f"无法安装 Office 加载项清单：{error}") from None
    return EnableResult(server, manifests, trusted_now)


def disable(server: OfficeAddinServer | None, *, home: Path | None = None) -> None:
    if server is not None:
        server.stop()
    sideload.remove_manifest(home=home)


def remove_trust(state: Path | None = None) -> None:
    certificate.untrust(certificate.certificate_paths(state or state_directory()))
