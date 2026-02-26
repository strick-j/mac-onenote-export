"""Tests for onenote_export.cli module."""

import logging
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from onenote_export.cli import main, _deduplicate_sections


TEST_DATA = Path(__file__).parent / "test_data"
NOTEBOOK_DIR = TEST_DATA / "Example-NoteBook-1"


class TestDeduplicateSections:
    """Tests for _deduplicate_sections."""

    def test_no_duplicates(self, tmp_path):
        files = [
            tmp_path / "Notes.one",
            tmp_path / "Tasks.one",
        ]
        for f in files:
            f.touch()
        result = _deduplicate_sections(files)
        assert len(result) == 2

    def test_keeps_latest_version(self, tmp_path):
        old = tmp_path / "ADI (On 10-3-22).one"
        new = tmp_path / "ADI (On 2-25-26).one"
        old.touch()
        new.touch()
        result = _deduplicate_sections([old, new])
        assert len(result) == 1
        assert result[0] == new

    def test_dotone_date_pattern(self, tmp_path):
        old = tmp_path / "ADP.one (On 8-24-22).one"
        new = tmp_path / "ADP.one (On 8-24-25).one"
        old.touch()
        new.touch()
        result = _deduplicate_sections([old, new])
        assert len(result) == 1
        assert result[0] == new

    def test_undated_file_kept_when_no_dated_version(self, tmp_path):
        f = tmp_path / "Notes.one"
        f.touch()
        result = _deduplicate_sections([f])
        assert len(result) == 1
        assert result[0] == f

    def test_dated_wins_over_undated(self, tmp_path):
        undated = tmp_path / "Notes.one"
        dated = tmp_path / "Notes (On 2-25-26).one"
        undated.touch()
        dated.touch()
        result = _deduplicate_sections([undated, dated])
        assert len(result) == 1
        assert result[0] == dated

    def test_empty_list(self):
        assert _deduplicate_sections([]) == []

    def test_two_digit_year_normalization(self, tmp_path):
        old = tmp_path / "Test (On 1-1-99).one"  # 1999
        new = tmp_path / "Test (On 1-1-24).one"  # 2024
        old.touch()
        new.touch()
        result = _deduplicate_sections([old, new])
        assert len(result) == 1
        assert result[0] == new


class TestMainHappyPath:
    """Tests for main() with real test data."""

    def test_successful_export(self, tmp_path):
        """main() returns 0 and creates output files."""
        result = main(["-i", str(NOTEBOOK_DIR), "-o", str(tmp_path / "out")])
        assert result == 0
        out_dir = tmp_path / "out"
        assert out_dir.exists()
        md_files = list(out_dir.rglob("*.md"))
        assert len(md_files) >= 1

    def test_flat_output(self, tmp_path):
        """--flat flag produces flat directory structure."""
        result = main([
            "-i", str(NOTEBOOK_DIR),
            "-o", str(tmp_path / "flat_out"),
            "--flat",
        ])
        assert result == 0
        out_dir = tmp_path / "flat_out"
        md_files = list(out_dir.rglob("*.md"))
        assert len(md_files) >= 1


class TestMainErrorHandling:
    """Tests for main() error paths."""

    def test_missing_input_dir(self, tmp_path):
        """Returns 1 for non-existent input directory."""
        result = main([
            "-i", str(tmp_path / "nonexistent"),
            "-o", str(tmp_path / "out"),
        ])
        assert result == 1

    def test_empty_input_dir(self, tmp_path):
        """Returns 1 when no .one files found."""
        empty_dir = tmp_path / "empty"
        empty_dir.mkdir()
        result = main(["-i", str(empty_dir), "-o", str(tmp_path / "out")])
        assert result == 1

    def test_parse_error_collected(self, tmp_path):
        """Parse errors are collected but don't crash the CLI."""
        # Create a fake .one file that will fail to parse
        bad_dir = tmp_path / "bad_notebook"
        bad_dir.mkdir()
        bad_file = bad_dir / "bad.one"
        bad_file.write_bytes(b"not a real onenote file")

        result = main(["-i", str(bad_dir), "-o", str(tmp_path / "out")])
        # Should return 2 (errors with no files written) or continue
        assert result in (0, 2)

    def test_converter_error_handled(self, tmp_path):
        """Converter errors are collected but don't crash."""
        with patch(
            "onenote_export.cli.MarkdownConverter"
        ) as mock_converter_cls:
            mock_converter = MagicMock()
            mock_converter.convert_notebook.side_effect = RuntimeError("write failed")
            mock_converter_cls.return_value = mock_converter

            result = main([
                "-i", str(NOTEBOOK_DIR),
                "-o", str(tmp_path / "out"),
            ])
            # Should handle the error gracefully
            assert result in (0, 2)


class TestMainLogging:
    """Tests for main() logging configuration."""

    def test_verbose_sets_info_level(self, tmp_path):
        """--verbose flag sets INFO logging."""
        with patch("onenote_export.cli.logging.basicConfig") as mock_config:
            main([
                "-i", str(NOTEBOOK_DIR),
                "-o", str(tmp_path / "out"),
                "--verbose",
            ])
            mock_config.assert_called_once()
            call_kwargs = mock_config.call_args[1]
            assert call_kwargs["level"] == logging.INFO

    def test_debug_sets_debug_level(self, tmp_path):
        """--debug flag sets DEBUG logging."""
        with patch("onenote_export.cli.logging.basicConfig") as mock_config:
            main([
                "-i", str(NOTEBOOK_DIR),
                "-o", str(tmp_path / "out"),
                "--debug",
            ])
            mock_config.assert_called_once()
            call_kwargs = mock_config.call_args[1]
            assert call_kwargs["level"] == logging.DEBUG

    def test_default_warning_level(self, tmp_path):
        """Default log level is WARNING."""
        with patch("onenote_export.cli.logging.basicConfig") as mock_config:
            main([
                "-i", str(NOTEBOOK_DIR),
                "-o", str(tmp_path / "out"),
            ])
            mock_config.assert_called_once()
            call_kwargs = mock_config.call_args[1]
            assert call_kwargs["level"] == logging.WARNING
