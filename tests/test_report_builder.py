from pathlib import Path
from xmlchecks.pkg.report_builder import build_report
from xmlchecks.pkg.crc import compute_crc32


def create_file(dir, name, content):
    p = dir / name
    p.write_text(content)
    return p


def test_build_report_scenarios(tmp_path):
    good = create_file(tmp_path, 'good.ifc', 'good')
    badcrc = create_file(tmp_path, 'badcrc.ifc', 'bad')
    name_mismatch = create_file(tmp_path, 'name_mismatch.ifc', 'same')
    extra = create_file(tmp_path, 'extra.ifc', 'extra')

    crc_good = f"{compute_crc32(good):08X}"
    crc_name = f"{compute_crc32(name_mismatch):08X}"

    xml_map = {
        'good.ifc': {'crc_hex': crc_good},
        'badcrc.ifc': {'crc_hex': '00000000'},
        'other.ifc': {'crc_hex': crc_name},
        'missing.ifc': {'crc_hex': '12345678'},
    }

    rows = build_report(xml_map, [good, badcrc, name_mismatch, extra])
    status = {row['Имя файла'] or row['Файл из XML']: row['Статус'] for row in rows}

    assert status['good.ifc'] == 'OK'
    assert status['badcrc.ifc'] == 'CRC_MISMATCH'
    assert status['name_mismatch.ifc'] == 'NAME_MISMATCH'
    assert status['extra.ifc'] == 'ERROR_IFC_EXTRA'
    assert status['missing.ifc'] == 'ERROR_XML_EXTRA'

    row_nm = next(r for r in rows if r['Имя файла'] == 'name_mismatch.ifc')
    assert row_nm['CRC-32 XML'] == crc_name
