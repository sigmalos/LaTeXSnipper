"""Locate the files served to Office: add-in pages, shared editor, MathJax.

Source checkouts read the Windows plugin's shared editor and renderer files in
place, so both platforms always run one implementation. Packaged builds carry
copies under ``application/office_addin`` (see
``scripts/build_application_resources.py``).
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from runtime.app_paths import get_app_root, resource_path

PACKAGED_DIRECTORY = "office_addin"
TYPOGRAPHY_ADAPTER = "MathJaxTypographyAdapter.js"
MANIFEST_TEMPLATE = "manifest.xml"


@dataclass(frozen=True)
class OfficeAddinAssets:
    web: Path
    shared_editor: Path
    typography_adapter: Path
    manifest_template: Path
    mathjax: Path
    support: Path

    def validate(self) -> None:
        required = (
            self.web / "taskpane.html",
            self.shared_editor / "editor.js",
            self.typography_adapter,
            self.manifest_template,
            self.mathjax / "startup.js",
        )
        missing = [str(path) for path in required if not path.is_file()]
        if missing:
            raise FileNotFoundError("Office 加载项资源缺失：" + ", ".join(missing))


def packaged_assets(root: Path, mathjax: Path) -> OfficeAddinAssets:
    return OfficeAddinAssets(
        web=root / "web",
        shared_editor=root / "shared",
        typography_adapter=root / TYPOGRAPHY_ADAPTER,
        manifest_template=root / MANIFEST_TEMPLATE,
        mathjax=mathjax,
        support=root / "support",
    )


def source_assets(repository: Path, mathjax: Path) -> OfficeAddinAssets:
    plugin = repository / "office_plugin"
    return OfficeAddinAssets(
        web=plugin / "macos/web",
        shared_editor=plugin / "src/LaTeXSnipper.OfficePlugin.Editor/EditorAssets",
        typography_adapter=plugin
        / "src/LaTeXSnipper.OfficePlugin.Rendering"
        / TYPOGRAPHY_ADAPTER,
        manifest_template=plugin / "macos" / MANIFEST_TEMPLATE,
        mathjax=mathjax,
        # Help-page images shared with the Windows plugin.
        support=plugin / "hosts/WordAddIn/EditorAssets/support",
    )


def resolve_assets() -> OfficeAddinAssets:
    mathjax = Path(resource_path("assets/MathJax"))
    packaged = get_app_root() / PACKAGED_DIRECTORY
    if packaged.is_dir():
        return packaged_assets(packaged, mathjax)
    return source_assets(get_app_root().parent, mathjax)
