#!/usr/bin/env python3
"""Generate an 800,000-cell XLSX and ask LibreOffice to save a real BIFF8 XLS.
All artifacts go to the host mount; the disposable LO profile uses /dev/shm.
"""
import hashlib
import json
import pathlib
import subprocess
import tempfile
import time
import zipfile

out = pathlib.Path('tmp/legacy/artifacts/C1-large').resolve()
out.mkdir(parents=True, exist_ok=True)
seed = out / 'large.xlsx'
main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
parts = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ''.join('<Override PartName="/xl/' + p + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.' + t + '+xml"/>' for p, t in [('workbook.xml', 'sheet.main'), ('styles.xml', 'styles'), ('worksheets/sheet1.xml', 'worksheet')]) + '</Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + rels + '/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + rels + '/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="' + rels + '/styles" Target="styles.xml"/></Relationships>',
    'xl/workbook.xml': '<workbook xmlns="' + main + '" xmlns:r="' + rels + '"><sheets><sheet name="Large" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/styles.xml': '<styleSheet xmlns="' + main + '"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
}
with zipfile.ZipFile(seed, 'w', zipfile.ZIP_STORED) as z:
    for name, value in parts.items(): z.writestr(name, value)
    with z.open('xl/worksheets/sheet1.xml', 'w') as f:
        f.write(('<worksheet xmlns="' + main + '"><dimension ref="A1:P50000"/><sheetData>').encode())
        for row in range(1, 50001):
            cells = ''.join(f'<c r="{chr(65 + col)}{row}" t="n"><v>{((row - 1) * 16 + col) / 8}</v></c>' for col in range(16))
            f.write(f'<row r="{row}">{cells}</row>'.encode())
        f.write(b'</sheetData></worksheet>')
start = time.monotonic()
with tempfile.TemporaryDirectory(prefix='i24-large-', dir='/dev/shm') as profile:
    result = subprocess.run(['soffice', '-env:UserInstallation=' + pathlib.Path(profile).as_uri(), '--headless', '--convert-to', 'xls:MS Excel 97', '--outdir', str(out / 'binary'), str(seed)], capture_output=True, timeout=240)
(out / 'libreoffice.log').write_bytes(result.stdout + result.stderr)
binary = out / 'binary/large.xls'
info = {'cells': 800000, 'rows': 50000, 'columns': 16, 'seedBytes': seed.stat().st_size, 'binaryBytes': binary.stat().st_size, 'loSeconds': time.monotonic() - start, 'sourceHash': hashlib.sha256(binary.read_bytes()).hexdigest()}
(out / 'source.json').write_text(json.dumps(info, indent=2) + '\n')
print(json.dumps(info))
