import sys
import tempfile
import unicodedata
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app


class WindowsZipMakerTests(unittest.TestCase):
    def test_safe_arcname_sanitizes_windows_incompatible_names(self):
        self.assertEqual(app.safe_arcname("a:b.txt"), "a_b.txt")
        self.assertEqual(app.safe_arcname("bad|name?.txt"), "bad_name_.txt")
        self.assertEqual(app.safe_arcname("CON.txt"), "_CON.txt")
        self.assertEqual(app.safe_arcname("folder/AUX"), "folder/_AUX")
        self.assertEqual(app.safe_arcname("../safe.txt"), "safe.txt")
        self.assertEqual(app.safe_arcname("C:/Users/example/file.txt"), "Users/example/file.txt")

    def test_build_zip_normalizes_names_and_skips_macos_metadata(self):
        with tempfile.TemporaryDirectory(prefix="wzm_test_") as tmp:
            root = Path(tmp)
            source = root / unicodedata.normalize("NFD", "테스트 폴더")
            source.mkdir()
            (source / unicodedata.normalize("NFD", "한글 파일.txt")).write_text("hello", encoding="utf-8")
            (source / "bad|name?.txt").write_text("bad", encoding="utf-8")
            (source / ".DS_Store").write_text("skip", encoding="utf-8")
            (source / "._한글 파일.txt").write_text("skip", encoding="utf-8")
            (source / "empty folder").mkdir()

            out_path, file_count = app.build_zip_from_paths(str(source), str(root / "out"), "결과", False)

            self.assertEqual(file_count, 2)
            with zipfile.ZipFile(out_path) as zf:
                names = zf.namelist()

            self.assertIn("테스트 폴더/", names)
            self.assertIn("테스트 폴더/한글 파일.txt", names)
            self.assertIn("테스트 폴더/bad_name_.txt", names)
            self.assertIn("테스트 폴더/empty folder/", names)
            self.assertFalse(any(".DS_Store" in name or "/._" in name for name in names))
            self.assertTrue(all(unicodedata.is_normalized("NFC", name) for name in names))


if __name__ == "__main__":
    unittest.main()
