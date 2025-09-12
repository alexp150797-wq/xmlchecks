from pathlib import Path
import zlib
from xmlchecks.pkg.crc import compute_crc32

def test_compute_crc32(tmp_path):
    p = tmp_path / 'data.bin'
    data = b'hello world'
    p.write_bytes(data)
    expected = zlib.crc32(data) & 0xFFFFFFFF
    assert compute_crc32(p) == expected
