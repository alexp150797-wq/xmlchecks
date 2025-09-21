from pathlib import Path

from xmlchecks import diagnostics_iul


def test_main_without_arguments(monkeypatch, capsys):
    calls = []

    monkeypatch.setattr(diagnostics_iul, "_print_heading", lambda *_: None)
    monkeypatch.setattr(diagnostics_iul, "_diagnose_pkg_import", lambda: None)
    monkeypatch.setattr(diagnostics_iul, "_check_tesseract_configuration", lambda *_: Path())
    monkeypatch.setattr(diagnostics_iul, "_list_tesseract_languages", lambda *_: None)
    monkeypatch.setattr(diagnostics_iul, "_check_required_modules", lambda: None)
    monkeypatch.setattr(diagnostics_iul, "_import_module", lambda *_: None)

    def fake_analyze(path):
        calls.append(path)

    monkeypatch.setattr(diagnostics_iul, "_analyze_pdf", fake_analyze)
    monkeypatch.setattr(diagnostics_iul.sys, "argv", ["diagnostics_iul.py"])

    diagnostics_iul.main()

    captured = capsys.readouterr()
    assert calls == []
    assert "Передайте путь к PDF" in captured.out
