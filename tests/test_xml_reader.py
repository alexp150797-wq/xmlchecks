from pathlib import Path
from xmlchecks.pkg.xml_reader import extract_from_xml, DEFAULT_RULES

def test_extract_from_xml(tmp_path):
    xml_content = '''<?xml version="1.0"?>\n<Root>\n  <ModelFile>\n    <FileName>file1.ifc</FileName>\n    <FileChecksum>ABCDEF12</FileChecksum>\n    <FileFormat>IFC</FileFormat>\n  </ModelFile>\n</Root>'''
    xml_path = tmp_path / 'data.xml'
    xml_path.write_text(xml_content, encoding='utf-8')
    res = extract_from_xml(xml_path, DEFAULT_RULES)
    assert res == {'file1.ifc': {'crc_hex': 'ABCDEF12', 'format': 'IFC'}}
