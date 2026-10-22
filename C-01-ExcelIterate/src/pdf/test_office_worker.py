"""Opt-in Office/WPS export smoke test, using only freshly created test documents.

Starts installed Office/WPS via COM. Never opens user documents or installs the add-in.
Run separately from builds: python src/pdf/test_office_worker.py
Test inspection requires pymupdf; fixtures use only Python's standard library.
"""
from pathlib import Path
import hashlib
import subprocess
import uuid
import xml.etree.ElementTree as ET
import zipfile
import pymupdf

ROOT = Path(__file__).resolve().parents[2]
folder = ROOT / 'build/pdf-office-tests' / uuid.uuid4().hex[:8]
folder.mkdir(parents=True)


def package(name, parts):
    path = folder / name
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for key, text in parts.items():
            archive.writestr(key, text)
    return path


doc = package('sample.docx', {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>PDF export Word sample 12345</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
})
book = package('sample.xlsx', {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sample" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="45" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>PDF export Excel sample 67890</t></is></c></row></sheetData></worksheet>',
})
before = [hashlib.sha256(p.read_bytes()).digest() for p in (doc, book)]
job = ET.Element('PdfJob', mode='convert', output=str(folder))
for path in (doc, book):
    ET.SubElement(job, 'File').text = str(path)
request = folder / 'request.xml'
ET.ElementTree(job).write(request, encoding='utf-8', xml_declaration=True)
print('Office test fixtures:', folder, flush=True)
result = subprocess.run([str(ROOT / 'build/pdf-worker/PdfWorker.exe'), str(request)], timeout=90)
report = ET.parse(str(request) + '.result').getroot()
assert result.returncode == 0, ET.tostring(report, encoding='unicode')
files = [Path(p.text) for p in report.findall('Success')]
assert len(files) == 2
for path, expected in zip(files, ['12345', '67890']):
    with pymupdf.open(path) as pdf:
        assert len(pdf) >= 1 and expected in ''.join(page.get_text() for page in pdf)
assert before == [hashlib.sha256(p.read_bytes()).digest() for p in (doc, book)]
print('PASS: real DOCX/XLSX exports; PDF text verified; both sources unchanged.')

# Exercise actual mixed Office/image conversion followed by all-or-nothing merge.
image = folder / 'sample-image.png'
pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 60, 40), False)
pix.clear_with(100)
pix.save(image)
merged_job = ET.Element('PdfJob', mode='convertmerge', output=str(folder))
for path in (doc, image, book):
    ET.SubElement(merged_job, 'File').text = str(path)
merged_request = folder / 'merged-request.xml'
ET.ElementTree(merged_job).write(merged_request, encoding='utf-8', xml_declaration=True)
merged_result = subprocess.run([str(ROOT / 'build/pdf-worker/PdfWorker.exe'), str(merged_request)], timeout=90)
merged_report = ET.parse(str(merged_request) + '.result').getroot()
assert merged_result.returncode == 0, ET.tostring(merged_report, encoding='unicode')
merged_files = [Path(node.text) for node in merged_report.findall('Success')]
assert len(merged_files) == 1
with pymupdf.open(merged_files[0]) as merged:
    assert len(merged) == 3
    assert '12345' in merged[0].get_text()
    assert len(merged[1].get_images()) == 1
    assert '67890' in merged[2].get_text()
assert before == [hashlib.sha256(p.read_bytes()).digest() for p in (doc, book)]
print('PASS: real mixed Word/image/Excel merge, ordering and source digests verified.')
