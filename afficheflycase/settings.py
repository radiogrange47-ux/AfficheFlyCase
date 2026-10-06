"""Persistent per-user settings for AfficheFlyCase."""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import asdict, dataclass
from pathlib import Path


@dataclass
class AppSettings:
    show: str = ""
    dates: str = ""
    venue: str = ""
    logo_path: str = ""


class SettingsError(OSError):
    """Raised when application settings cannot be read or saved."""


def settings_directory() -> Path:
    app_data = os.environ.get("APPDATA")
    base = Path(app_data) if app_data else Path.home() / "AppData" / "Roaming"
    return base / "AfficheFlyCase"


def load_settings() -> AppSettings:
    path = settings_directory() / "settings.json"
    if not path.exists():
        return AppSettings()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return AppSettings(
            show=str(data.get("show", "")),
            dates=str(data.get("dates", "")),
            venue=str(data.get("venue", "")),
            logo_path=str(data.get("logo_path", "")),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, AttributeError, TypeError) as error:
        raise SettingsError(f"Impossible de lire les préférences AfficheFlyCase : {error}") from error


def save_settings(settings: AppSettings) -> None:
    directory = settings_directory()
    temporary_path: Path | None = None
    try:
        directory.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            prefix=".settings_",
            suffix=".json",
            dir=directory,
            delete=False,
        ) as stream:
            temporary_path = Path(stream.name)
            json.dump(asdict(settings), stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        os.replace(temporary_path, directory / "settings.json")
        temporary_path = None
    except (OSError, TypeError, ValueError) as error:
        raise SettingsError(f"Impossible d’enregistrer les préférences AfficheFlyCase : {error}") from error
    finally:
        if temporary_path and temporary_path.exists():
            temporary_path.unlink()


def store_logo(source: str | Path) -> Path:
    """Validate and copy a selected image into the persistent app data directory."""
    from reportlab.lib.utils import ImageReader

    source_path = Path(source)
    if source_path.suffix.casefold() not in {".png", ".jpg", ".jpeg"}:
        raise SettingsError("Format de logo non pris en charge. Choisissez une image PNG ou JPEG.")
    if not source_path.is_file():
        raise SettingsError("Le fichier du logo est introuvable.")
    try:
        ImageReader(str(source_path)).getSize()
        directory = settings_directory()
        directory.mkdir(parents=True, exist_ok=True)
        destination = directory / f"logo{source_path.suffix.casefold()}"
        if source_path.resolve() != destination.resolve():
            from shutil import copy2

            copy2(source_path, destination)
        return destination
    except (OSError, ValueError, TypeError) as error:
        raise SettingsError(f"Impossible d’utiliser ou de mémoriser le logo : {error}") from error
