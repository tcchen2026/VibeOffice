/* A name another workbook defines on one of its sheets ('[1]Sheet'!Name), calculated from the values cached with
   this file, as Excel does: each sheet's definition is its own. Such names were #REF!, and the second sheet's
   definition replaced the first's. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const L = require('./load.js');

const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const workbook = async () => {
  const ct = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="${ct}sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${ct}worksheet+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="${ct}externalLink+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId2"/></externalReferences><definedNames><definedName name="Cost">Sheet1!$A$1</definedName></definedNames></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1"><v>41</v></c></row><row r="3"><c r="E3"><f>Cost*'[1]MarkupSheet'!Markup_Cost</f><v>0</v></c></row><row r="4"><c r="E4"><f>Cost*'[1]MarkupTestSheet'!Markup_Cost</f><v>0</v></c></row><row r="5"><c r="E5"><f>Cost*[1]!Markup_All</f><v>0</v></c></row></sheetData></worksheet>`,
    'xl/externalLinks/externalLink1.xml': `<externalLink xmlns="${S}" xmlns:r="${R}"><externalBook r:id="rId1"><sheetNames><sheetName val="MarkupSheet"/><sheetName val="MarkupTestSheet"/></sheetNames><definedNames><definedName name="Markup_All" refersTo="=MarkupSheet!$B$2"/><definedName name="Markup_Cost" refersTo="=MarkupSheet!$B$1" sheetId="0"/><definedName name="Markup_Cost" refersTo="=MarkupTestSheet!$B$1" sheetId="1"/></definedNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="B1"><v>1.08</v></cell></row><row r="2"><cell r="B2"><v>2</v></cell></row></sheetData><sheetData sheetId="1"><row r="1"><cell r="B1"><v>1.14</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>`,
    'xl/externalLinks/_rels/externalLink1.xml.rels': `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/externalLinkPath" Target="markup.xls" TargetMode="External"/></Relationships>`,
  };
  const blob = await L.zip.write(Object.entries(parts).map(([name, data]) => ({ name, data })));
  return L.xlsxRead.read(new Uint8Array(await blob.arrayBuffer()));
};

test('a name defined on one sheet of another workbook takes that sheet\'s definition', async () => {
  const wb = await workbook();
  L.calc.recalcAll(wb);
  const at = (a) => { const p = L.formula.parseCell(a), c = wb.sheets[0].get(p.r, p.c); return c && (L.model.isErr(c.v) ? c.v.e : c.v); };
  assert.ok(Math.abs(at('E3') - 44.28) < 1e-9, `MarkupSheet's Markup_Cost: ${at('E3')}`);
  assert.ok(Math.abs(at('E4') - 46.74) < 1e-9, `MarkupTestSheet's Markup_Cost: ${at('E4')}`);
  assert.equal(at('E5'), 82, 'the workbook\'s own name, as before');
});
