import os
from xmlchecks.pkg.report_builder_iul import build_report_iul, _fmt_mtime
from xmlchecks.pkg.crc import compute_crc32
from xmlchecks.pkg.iul_reader import IulEntry


def create_file(dir, name, content, mtime):
    p = dir / name
    p.write_text(content)
    os.utime(p, (mtime, mtime))
    return p


def test_build_report_iul_scenarios(tmp_path):
    base = 1700000000  # fixed timestamp
    good = create_file(tmp_path, 'good.ifc', 'good', base)
    crc_bad = create_file(tmp_path, 'crc_bad.ifc', 'crc', base + 1)
    size_bad = create_file(tmp_path, 'size_bad.ifc', 'size', base + 2)
    dt_bad = create_file(tmp_path, 'dt_bad.ifc', 'dt', base + 3)
    pdf_bad = create_file(tmp_path, 'pdf_bad.ifc', 'pdf', base + 4)
    name_mismatch_file = create_file(tmp_path, 'name_mismatch.ifc', 'same', base + 5)
    extra = create_file(tmp_path, 'extra.ifc', 'extra', base + 6)

    def info(p):
        return f"{compute_crc32(p):08X}", _fmt_mtime(p.stat().st_mtime), p.stat().st_size

    crc_good, dt_good, size_good = info(good)
    crc_crc_bad, dt_crc, size_crc = info(crc_bad)
    crc_size, dt_size, size_size = info(size_bad)
    crc_dt, dt_dt, size_dt = info(dt_bad)
    crc_pdf, dt_pdf, size_pdf = info(pdf_bad)
    crc_name, dt_name, size_name = info(name_mismatch_file)

    iul_map = {
        'good.ifc': IulEntry('good.ifc', crc_good, dt_good, size_good, 'ctx', 'good_ИУЛ.pdf'),
        'crc_bad.ifc': IulEntry('crc_bad.ifc', 'FFFFFFFF', dt_crc, size_crc, 'ctx', 'crc_bad_ИУЛ.pdf'),
        'size_bad.ifc': IulEntry('size_bad.ifc', crc_size, dt_size, size_size + 1, 'ctx', 'size_bad_ИУЛ.pdf'),
        'dt_bad.ifc': IulEntry('dt_bad.ifc', crc_dt, '01.01.2000 00:00', size_dt, 'ctx', 'dt_bad_ИУЛ.pdf'),
        'pdf_bad.ifc': IulEntry('pdf_bad.ifc', crc_pdf, dt_pdf, size_pdf, 'ctx', 'pdf_bad.pdf'),
        'other.ifc': IulEntry('other.ifc', crc_name, dt_name, size_name, 'ctx', 'name_mismatch_ИУЛ.pdf'),
        'missing.ifc': IulEntry('missing.ifc', 'ABCDEF12', '01.01.2024 00:00', 123, 'ctx', 'missing_ИУЛ.pdf'),
    }

    rows = build_report_iul(iul_map, [good, crc_bad, size_bad, dt_bad, pdf_bad, name_mismatch_file, extra])
    status = {row['Имя файла'] or row['Файл из ИУЛ']: row['Статус'] for row in rows}

    assert status['good.ifc'] == 'OK'
    assert status['crc_bad.ifc'] == 'CRC_MISMATCH'
    assert status['size_bad.ifc'] == 'SIZE_MISMATCH'
    assert status['dt_bad.ifc'] == 'DT_MISMATCH'
    assert status['pdf_bad.ifc'] == 'PDF_NAME_MISMATCH'
    assert status['name_mismatch.ifc'] == 'NAME_MISMATCH'
    assert status['extra.ifc'] == 'ERROR_IFC_EXTRA'
    assert status['missing.ifc'] == 'ERROR_IUL_EXTRA'
