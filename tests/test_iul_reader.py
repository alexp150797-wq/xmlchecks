from pathlib import Path
from xmlchecks.pkg.iul_reader import extract_iul_entries_from_pdf

def test_extract_iul_entries_from_pdf(monkeypatch, tmp_path):
    pdf_path = tmp_path / 'doc.pdf'
    pdf_path.write_bytes(b'%PDF-1.4')
    sample_text = 'CRC-32 ABCDEF12\nfile1.ifc 01.02.2024 12:34 1234'
    monkeypatch.setattr('xmlchecks.pkg.iul_reader._extract_text_pypdf2', lambda p: sample_text)
    monkeypatch.setattr('xmlchecks.pkg.iul_reader._extract_text_ocr', lambda p: '')
    entries = extract_iul_entries_from_pdf(pdf_path)
    assert len(entries) == 1
    e = entries[0]
    assert e.basename == 'file1.ifc'
    assert e.crc_hex == 'ABCDEF12'
    assert e.dt_str == '01.02.2024 12:34'
    assert e.size_bytes == 1234
    assert e.source_pdf == pdf_path.name


def test_extract_iul_with_spaces(monkeypatch, tmp_path):
    pdf_path = tmp_path / 'doc.pdf'
    pdf_path.write_bytes(b'%PDF-1.4')
    sample_text = 'CRC-32 ABCDEF12\n1-2024-60_П_ТКР_ОХ.П_Пролетное строение.ifc 01.02.2024 12:34 1234'
    monkeypatch.setattr('xmlchecks.pkg.iul_reader._extract_text_pypdf2', lambda p: sample_text)
    monkeypatch.setattr('xmlchecks.pkg.iul_reader._extract_text_ocr', lambda p: '')
    entries = extract_iul_entries_from_pdf(pdf_path)
    assert entries[0].basename == '1-2024-60_П_ТКР_ОХ.П_Пролетное строение.ifc'


def test_extract_iul_uses_ocr_fallback(monkeypatch, tmp_path):
    pdf_path = tmp_path / 'doc.pdf'
    pdf_path.write_bytes(b'%PDF-1.4')
    monkeypatch.setattr('xmlchecks.pkg.iul_reader._extract_text_pypdf2', lambda p: '')

    sample_text = 'CRC-32 12345678\nscan.ifc 11.03.2024 10:10 98765'
    called = {}

    def fake_ocr(path, dpi=300):
        called['dpi'] = dpi
        return sample_text

    monkeypatch.setattr('xmlchecks.pkg.iul_reader._extract_text_ocr', fake_ocr)

    entries = extract_iul_entries_from_pdf(pdf_path)
    assert entries and entries[0].basename == 'scan.ifc'
    assert entries[0].crc_hex == '12345678'
    assert called.get('dpi') == 300
