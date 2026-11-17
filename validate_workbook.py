"""Verify Excel workbook ZIP structure and readable worksheet XML."""
from pathlib import Path
from zipfile import ZipFile, BadZipFile
from xml.etree import ElementTree as ET

path = Path(__file__).with_name("Project Tracking.xlsx")
assert path.is_file(), "Missing Excel workbook"
try:
    with ZipFile(path) as z:
        assert z.testzip() is None, "Corrupt workbook archive member"
        names = set(z.namelist())
        assert "[Content_Types].xml" in names and "xl/workbook.xml" in names
        sheets = [n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        assert sheets, "No worksheet XML found"
        for name in ["xl/workbook.xml", *sheets]:
            ET.fromstring(z.read(name))
except BadZipFile as e:
    raise AssertionError("Workbook is not a valid XLSX archive") from e
print(f"Workbook structure passed: {len(sheets)} worksheet XML file(s).")
