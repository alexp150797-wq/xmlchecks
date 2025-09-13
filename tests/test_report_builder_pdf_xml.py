from xmlchecks.pkg.report_builder_pdf_xml import build_report_pdf_xml
from xmlchecks.pkg.crc import compute_crc32


def test_crc_field_filled_on_name_mismatch(tmp_path):
    pdf = tmp_path / "sample.pdf"
    pdf.write_text("content")
    crc = f"{compute_crc32(pdf):08X}"
    xml_map = {"other.pdf": {"crc_hex": crc}}
    rows = build_report_pdf_xml(xml_map, [pdf], case_sensitive=True)
    row = rows[0]
    assert row["Статус"] == "NAME_MISMATCH"
    assert row["CRC-32 XML"] == crc
    assert row["Файл из XML"] == "other.pdf"
