/**
 * Create a synthetic DOCX fixture shaped like the tables AI document
 * generators emit (docx npm / python-docx): auto-width tables whose
 * w:tblGrid carries tiny placeholder widths (100 twips per column — the
 * docx npm default when no columnWidths are given), and percent widths
 * written in the "%"-suffixed ST_MeasurementOrPercent form ("100%")
 * rather than fiftieths-of-a-percent (5000). Word autofits/expands both;
 * the editor must not render them as a few-pixel sliver.
 *
 * Run: bun scripts/create-ai-generated-table-fixture.mjs
 */

import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'e2e/fixtures/ai-generated-table-widths.docx');
const ZIP_DATE = new Date('2026-01-01T00:00:00Z');

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
  </w:style>
</w:styles>`;

const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>AI-Generated Table Widths Fixture</dc:title>
  <dc:creator>docx-editor fixture generator</dc:creator>
  <cp:lastModifiedBy>docx-editor fixture generator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

function p(text, options = {}) {
  const bold = options.bold ? '<w:b/>' : '';
  const size = options.size ?? 22;
  return `<w:p>
    <w:pPr><w:spacing w:after="${options.after ?? 120}" w:line="276" w:lineRule="auto"/></w:pPr>
    <w:r><w:rPr>${bold}<w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>
  </w:p>`;
}

const borders = `<w:tblBorders>
  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
</w:tblBorders>`;

const cellP = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

// Table 1 — the `docx` npm default shape: tblW auto/0, and a tblGrid whose
// every gridCol is the library's 100-twip placeholder. No w:tcW on cells.
const tcBare = (text) => `<w:tc>${cellP(text)}</w:tc>`;
const rowBare = (cells) => `<w:tr>${cells.map(tcBare).join('')}</w:tr>`;
const autofitPlaceholderTable = `<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    ${borders}
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>
  ${rowBare(['Feature', 'Plan A', 'Plan B'])}
  ${rowBare(['Storage included', '10 GB', '100 GB'])}
  ${rowBare(['Priority support', 'No', 'Yes, around the clock'])}
</w:tbl>`;

// Table 2 — "%"-suffixed percent widths (ST_MeasurementOrPercent):
// tblW "100%" pct, per-cell tcW "50%"/"25%" pct, bare gridCols.
const tcPct = (text, pct) => `<w:tc>
  <w:tcPr><w:tcW w:w="${pct}" w:type="pct"/></w:tcPr>
  ${cellP(text)}
</w:tc>`;
const percentTable = `<w:tbl>
  <w:tblPr>
    <w:tblW w:w="100%" w:type="pct"/>
    ${borders}
  </w:tblPr>
  <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
  <w:tr>${tcPct('Criterion', '50%')}${tcPct('Assessment', '50%')}</w:tr>
  <w:tr>${tcPct('Rendering fidelity', '50%')}${tcPct('Must match Word output', '50%')}</w:tr>
</w:tbl>`;

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${p('Comparison table', { bold: true, size: 32, after: 240 })}
    ${autofitPlaceholderTable}
    ${p('Percent-width table below.', { after: 240 })}
    ${percentTable}
    ${p('Closing generated paragraph.')}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const zip = new JSZip();
const opts = { date: ZIP_DATE, createFolders: false };
zip.file('[Content_Types].xml', contentTypesXml, opts);
zip.file('_rels/.rels', relsXml, opts);
zip.file('word/_rels/document.xml.rels', documentRelsXml, opts);
zip.file('word/document.xml', documentXml, opts);
zip.file('word/styles.xml', stylesXml, opts);
zip.file('docProps/core.xml', coreXml, opts);

const buffer = await zip.generateAsync({
  type: 'nodebuffer',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});
fs.writeFileSync(OUT, buffer);
console.log(`Created ${OUT}`);
