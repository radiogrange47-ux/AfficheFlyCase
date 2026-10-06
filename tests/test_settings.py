import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from afficheflycase.settings import (
    AppSettings,
    SettingsError,
    load_settings,
    save_settings,
    settings_directory,
    store_logo,
)


class AppSettingsTests(unittest.TestCase):
    def test_saves_and_restores_production_and_logo_preferences(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict("os.environ", {"APPDATA": directory}):
            expected = AppSettings("Le Show", "10–12 octobre", "Paris", str(Path(directory) / "logo.png"))
            save_settings(expected)
            actual = load_settings()
        self.assertEqual(actual, expected)

    def test_copies_logo_to_persistent_settings_directory(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict("os.environ", {"APPDATA": directory}):
            source = Path(directory) / "source-logo.png"
            Image.new("RGB", (24, 12), "red").save(source)
            stored = store_logo(source)
            source.unlink()
            self.assertTrue(stored.is_file())
            self.assertEqual(stored.parent, settings_directory())

    def test_rejects_non_image_logo_extension(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict("os.environ", {"APPDATA": directory}):
            image_path = Path(directory) / "logo.txt"
            image_path.write_text("not an image", encoding="utf-8")
            with self.assertRaisesRegex(SettingsError, "PNG ou JPEG"):
                store_logo(image_path)

    def test_rejects_corrupted_logo_image(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict("os.environ", {"APPDATA": directory}):
            image_path = Path(directory) / "logo.png"
            image_path.write_bytes(b"not a real image")
            with self.assertRaisesRegex(SettingsError, "Impossible d’utiliser ou de mémoriser"):
                store_logo(image_path)


if __name__ == "__main__":
    unittest.main()
