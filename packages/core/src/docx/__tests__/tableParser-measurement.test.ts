/**
 * parseTableMeasurement — ST_MeasurementOrPercent string forms.
 *
 * `w:w` on CT_TblWidth (w:tblW / w:tcW / w:tblInd / cell margins) admits three
 * value forms: an unqualified number (twips for dxa, fiftieths-of-a-percent
 * for pct), a `%`-suffixed percentage ("100%"), and a universal measure
 * ("0.5in"). Generators used by AI tooling (`docx` npm) emit the `%` form;
 * naive parseInt read "100%" as 100 fiftieths = 2% and collapsed the table
 * to a sliver.
 */
import { describe, test, expect } from 'bun:test';
import { parseXml, type XmlElement } from '../xmlParser';
import { parseWidth } from '../tableParser/properties';

function el(attrs: string): XmlElement {
  const root = parseXml(
    `<w:tblW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${attrs}/>`
  ) as unknown as { elements: XmlElement[] };
  return root.elements[0];
}

describe('parseTableMeasurement — percent string form', () => {
  test('"100%" pct → 5000 fiftieths', () => {
    expect(parseWidth(el('w:w="100%" w:type="pct"'))).toEqual({ value: 5000, type: 'pct' });
  });

  test('"50%" pct → 2500 fiftieths', () => {
    expect(parseWidth(el('w:w="50%" w:type="pct"'))).toEqual({ value: 2500, type: 'pct' });
  });

  test('decimal percentage "33.3%" rounds to fiftieths', () => {
    expect(parseWidth(el('w:w="33.3%" w:type="pct"'))).toEqual({ value: 1665, type: 'pct' });
  });

  test('the % value form wins over a mismatched declared type', () => {
    expect(parseWidth(el('w:w="100%" w:type="dxa"'))).toEqual({ value: 5000, type: 'pct' });
  });
});

describe('parseTableMeasurement — universal measure form', () => {
  test('"0.5in" → 720 twips dxa', () => {
    expect(parseWidth(el('w:w="0.5in" w:type="dxa"'))).toEqual({ value: 720, type: 'dxa' });
  });

  test('"12pt" → 240 twips dxa', () => {
    expect(parseWidth(el('w:w="12pt"'))).toEqual({ value: 240, type: 'dxa' });
  });

  test('"2.54cm" → 1440 twips dxa', () => {
    expect(parseWidth(el('w:w="2.54cm"'))).toEqual({ value: 1440, type: 'dxa' });
  });
});

describe('parseTableMeasurement — unqualified numbers stay as before', () => {
  test('pct fiftieths untouched (2500 = 50%)', () => {
    expect(parseWidth(el('w:w="2500" w:type="pct"'))).toEqual({ value: 2500, type: 'pct' });
  });

  test('dxa twips untouched', () => {
    expect(parseWidth(el('w:w="9360" w:type="dxa"'))).toEqual({ value: 9360, type: 'dxa' });
  });

  test('auto zero untouched', () => {
    expect(parseWidth(el('w:w="0" w:type="auto"'))).toEqual({ value: 0, type: 'auto' });
  });

  test('missing w defaults to 0', () => {
    expect(parseWidth(el('w:type="dxa"'))).toEqual({ value: 0, type: 'dxa' });
  });
});
