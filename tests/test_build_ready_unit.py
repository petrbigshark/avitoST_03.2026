import importlib.util
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import requests


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "build_ready_file.py"


def load_module():
    spec = importlib.util.spec_from_file_location("build_ready_file_unit", SCRIPT_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class BuildReadyUnitTests(unittest.TestCase):
    def test_normalize_id(self):
        mod = load_module()
        self.assertEqual(mod.normalize_id(10103001.0), "10103001")
        self.assertEqual(mod.normalize_id("10103001.0"), "10103001")
        self.assertEqual(mod.normalize_id("  10103001 "), "10103001")
        self.assertEqual(mod.normalize_id(None), "")

    def test_format_size_value(self):
        mod = load_module()
        self.assertEqual(mod.format_size_value(39.0), "39")
        self.assertEqual(mod.format_size_value(37.5), "37,5")
        self.assertEqual(mod.format_size_value("38.5"), "38,5")
        self.assertEqual(mod.format_size_value(None), "")

    def test_normalize_human_case(self):
        mod = load_module()
        self.assertEqual(mod.normalize_human_case("adidas Y-3 GAZELLE white"), "Adidas Y-3 Gazelle White")
        self.assertEqual(mod.normalize_human_case("WMNS GAZELLE INDOOR 'BLUE FUSION GUM'"), "Wmns Gazelle Indoor 'Blue Fusion Gum'")
        self.assertEqual(mod.normalize_human_case(None), "")

    def test_normalize_brand_offwhite_collab(self):
        mod = load_module()
        self.assertEqual(mod.normalize_brand("Nike X Off-White"), "Nike")
        self.assertEqual(mod.normalize_brand("OFF-WHITE x NIKE"), "Nike")
        self.assertEqual(mod.normalize_brand("aDIDAS"), "Adidas")

    def test_add_original_suffix(self):
        mod = load_module()
        self.assertEqual(mod.add_original_suffix("Adidas Samba"), "Adidas Samba Оригинал")
        self.assertEqual(mod.add_original_suffix("Adidas Samba Оригинал"), "Adidas Samba Оригинал")
        self.assertEqual(mod.add_original_suffix(""), "")

    def test_sanitize_multi_title(self):
        mod = load_module()
        raw = "Adidas Samba Og “White Black Gum” Оригинал"
        self.assertEqual(mod.sanitize_multi_title(raw), "Adidas Samba Og White Black Gum Оригинал")

    def test_mask_sensitive_text(self):
        mod = load_module()
        text = "400 Client Error for url: https://api.imgbb.com/1/upload?key=abc123secret"
        masked = mod.mask_sensitive_text(text)
        self.assertIn("key=***", masked)
        self.assertNotIn("abc123secret", masked)

    def test_choose_next_version_with_folder_and_legacy_files(self):
        mod = load_module()
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_dir = Path(tmp_dir)
            (out_dir / "Готовые файлы 01.03.2026 v2").mkdir()
            (out_dir / "Готовый файл 01.03.2026 v4.xlsx").write_text("", encoding="utf-8")
            (out_dir / "random.txt").write_text("", encoding="utf-8")
            mod.OUTPUT_ROOT_DIR = out_dir

            next_ver = mod.choose_next_version("01.03.2026")
            self.assertEqual(next_ver, 5)

    def test_parse_version_arg(self):
        mod = load_module()
        self.assertEqual(mod.parse_version_arg("v1"), 1)
        self.assertEqual(mod.parse_version_arg("12"), 12)
        with self.assertRaises(ValueError):
            mod.parse_version_arg("v0")
        with self.assertRaises(ValueError):
            mod.parse_version_arg("abc")

    def test_uploader_fail_fast_after_consecutive_errors(self):
        mod = load_module()
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            article_dir = root / "ART-FAIL"
            article_dir.mkdir(parents=True, exist_ok=True)
            for i in range(1, 12):
                (article_dir / f"{i}.jpg").write_bytes(b"jpg")

            uploader = mod.ImgbbUploader(
                api_key="dummy",
                photos_root=root,
                skip_upload=False,
                max_consecutive_errors=10,
            )

            def always_fail(_):
                raise RuntimeError("network-down")

            uploader._upload_image = always_fail  # type: ignore[attr-defined]

            with self.assertRaises(RuntimeError):
                uploader.links_for_article("ART-FAIL")

    def test_uploader_no_fail_fast_when_disabled(self):
        mod = load_module()
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            article_dir = root / "ART-OK"
            article_dir.mkdir(parents=True, exist_ok=True)
            for i in range(1, 4):
                (article_dir / f"{i}.jpg").write_bytes(b"jpg")

            uploader = mod.ImgbbUploader(
                api_key="dummy",
                photos_root=root,
                skip_upload=False,
                max_consecutive_errors=0,
            )

            def always_fail(_):
                raise RuntimeError("network-down")

            uploader._upload_image = always_fail  # type: ignore[attr-defined]
            result = uploader.links_for_article("ART-OK")
            self.assertIsNone(result)

    def test_heic_is_not_converted_by_default(self):
        mod = load_module()
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            heic_file = root / "sample.heic"
            heic_file.write_bytes(b"heic")
            uploader = mod.ImgbbUploader(
                api_key="dummy",
                photos_root=root,
                skip_upload=False,
                convert_heic_to_jpeg=False,
            )
            prepared = uploader._prepare_image(heic_file, root)  # type: ignore[attr-defined]
            self.assertEqual(prepared, heic_file)

    def test_heic_conversion_uses_best_quality(self):
        mod = load_module()
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            heic_file = root / "sample.heic"
            heic_file.write_bytes(b"heic")
            uploader = mod.ImgbbUploader(
                api_key="dummy",
                photos_root=root,
                skip_upload=False,
                convert_heic_to_jpeg=True,
            )
            with patch.object(mod.subprocess, "run", return_value=SimpleNamespace(returncode=0, stderr="", stdout="")) as run_mock:
                uploader._convert_heic_to_jpeg(heic_file, root)  # type: ignore[attr-defined]
                command = run_mock.call_args.args[0]
                self.assertIn("formatOptions", command)
                self.assertIn("best", command)

    def test_heic_fallback_to_jpeg_on_http_400(self):
        mod = load_module()
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            article_dir = root / "ART-HEIC"
            article_dir.mkdir(parents=True, exist_ok=True)
            heic_file = article_dir / "1.heic"
            heic_file.write_bytes(b"heic")
            jpg_file = root / "converted.jpg"
            jpg_file.write_bytes(b"jpg")

            uploader = mod.ImgbbUploader(
                api_key="dummy",
                photos_root=root,
                skip_upload=False,
                convert_heic_to_jpeg=False,
            )

            uploader._convert_heic_to_jpeg = lambda *_: jpg_file  # type: ignore[attr-defined]

            call_state = {"count": 0}

            def fake_upload(path: Path):
                call_state["count"] += 1
                if call_state["count"] == 1:
                    response = requests.Response()
                    response.status_code = 400
                    raise requests.HTTPError("400 Client Error", response=response)
                return f"https://img.test/{path.name}"

            uploader._upload_image = fake_upload  # type: ignore[attr-defined]
            links = uploader.links_for_article("ART-HEIC")
            self.assertEqual(links, "https://img.test/converted.jpg")
            self.assertEqual(call_state["count"], 2)

    def test_newlink_mode_decorates_url(self):
        mod = load_module()
        uploader = mod.ImgbbUploader(
            api_key="dummy",
            photos_root=Path("."),
            skip_upload=True,
            force_new_links=True,
            newlink_tag="01.03.2026-v1",
        )
        decorated = uploader._decorate_link_for_newlink_mode(
            "https://i.ibb.co/demo/image.jpg",
            article="B75806",
            index=1,
        )
        self.assertIn("newlink=", decorated)
        self.assertTrue(decorated.startswith("https://i.ibb.co/demo/image.jpg?"))

    def test_build_description(self):
        mod = load_module()
        text = mod.build_description(
            title="Nike SB Dunk Low",
            article="DA9658-500",
            size_lines=["• US 10 / EU 44 / RU 43 / 28 см", "• US 11 / EU 45 / RU 44 / 29 см"],
        )
        self.assertIn("Nike SB Dunk Low, DA9658-500", text)
        self.assertIn("Размеры :", text)
        self.assertIn("US 10 / EU 44 / RU 43 / 28 см", text)
        self.assertIn("📩 Напишите в личные сообщения", text)

    def test_kids_size_field(self):
        mod = load_module()
        self.assertEqual(mod.build_size_field("26", "16,5", kids=True), "26 (16,5 см)")
        self.assertEqual(mod.build_size_field("39", "25", kids=False), "39")
        self.assertEqual(mod.build_size_field("", "25", kids=True), "!НЕТ!")


if __name__ == "__main__":
    unittest.main()
