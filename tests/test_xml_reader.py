from pathlib import Path
from xmlchecks.pkg.xml_reader import extract_from_xml, DEFAULT_RULES

def test_extract_from_xml(tmp_path):
    xml_content = '''<?xml version="1.0"?>\n<Root>\n  <ModelFile>\n    <FileName>file1.ifc</FileName>\n    <FileChecksum>ABCDEF12</FileChecksum>\n    <FileFormat>IFC</FileFormat>\n  </ModelFile>\n</Root>'''
    xml_path = tmp_path / 'data.xml'
    xml_path.write_text(xml_content, encoding='utf-8')
    res = extract_from_xml(xml_path, DEFAULT_RULES)
    assert res == {'file1.ifc': {'crc_hex': 'ABCDEF12', 'format': 'IFC'}}


def test_extract_from_xml_sign_files_for_pdf(tmp_path):
    xml_content = '''<?xml version="1.0"?>
<Root>
  <ModelFile>
    <FileName>model.ifc</FileName>
    <FileChecksum>11111111</FileChecksum>
    <FileFormat>IFC</FileFormat>
    <SignFile>
      <FileName>report.pdf</FileName>
      <FileChecksum>22222222</FileChecksum>
      <FileFormat>PDF</FileFormat>
    </SignFile>
    <SignFile>
      <FileName>ignored.sig</FileName>
      <FileChecksum>33333333</FileChecksum>
      <FileFormat>SIG</FileFormat>
    </SignFile>
  </ModelFile>
</Root>'''
    xml_path = tmp_path / 'data_pdf.xml'
    xml_path.write_text(xml_content, encoding='utf-8')
    rules = DEFAULT_RULES.copy()
    rules["filter_format"] = "PDF"
    res = extract_from_xml(xml_path, rules, case_sensitive=True)
    assert res == {'report.pdf': {'crc_hex': '22222222', 'format': 'PDF'}}
