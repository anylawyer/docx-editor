/**
 * AI-generated table shapes must not collapse to a sliver.
 *
 * Document generators that AI tooling uses emit tables Word autofits on
 * open: `docx` npm defaults every `w:gridCol` to a 100-twip placeholder
 * when no column widths are given, and writes percent widths in the
 * `%`-suffixed ST_MeasurementOrPercent form ("100%"). Both used to render
 * as a few-pixel-wide unreadable column stack (chars wrapped one per line).
 *
 * Exercises the real chain: XML → parseTable → toProseDoc → toFlowBlocks →
 * resolveTableColumnWidths.
 */
import { afterAll, beforeAll, describe, test, expect } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { parseXml, type XmlElement } from '../../docx/xmlParser';
import { parseTable } from '../../docx/tableParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../toFlowBlocks';
import { resolveTableColumnWidths } from '../tableWidthUtils';
import type { Document } from '../../types/document';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const CONTENT_WIDTH_PX = 624; // Letter, 1in margins → 9360 twips

const cell = (txt: string, tcw = '') =>
  `<w:tc><w:tcPr>${tcw}</w:tcPr><w:p><w:r><w:t>${txt}</w:t></w:r></w:p></w:tc>`;

function tableXml(opts: { tblW?: string; grid?: string; tcw?: string; layout?: string }) {
  const row = (a: string, b: string) => `<w:tr>${cell(a, opts.tcw)}${cell(b, opts.tcw)}</w:tr>`;
  return `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:tblPr>${opts.tblW ?? ''}${opts.layout ?? ''}</w:tblPr>
    ${opts.grid ?? ''}
    ${row('Feature', 'Value')}
    ${row('Speed of light propagation', 'Very fast indeed, measured precisely')}
  </w:tbl>`;
}

function resolveWidths(xml: string): number[] {
  const root = parseXml(xml) as unknown as { elements: Array<XmlElement & { name?: string }> };
  const el = root.elements.find((e) => e.name === 'w:tbl');
  if (!el) throw new Error('no w:tbl element');
  const table = parseTable(el, null, null, null, null, null);
  const doc: Document = { package: { document: { content: [table] } } };
  const blocks = toFlowBlocks(toProseDoc(doc));
  const tb = blocks.find((b) => b.kind === 'table');
  if (!tb || tb.kind !== 'table') throw new Error('no table block');
  return resolveTableColumnWidths(tb, CONTENT_WIDTH_PX);
}

const total = (widths: number[]) => widths.reduce((a, b) => a + b, 0);

describe('AI-generated table width shapes', () => {
  test('docx npm default: tblW auto 0 + placeholder gridCol w=100 → full width', () => {
    const widths = resolveWidths(
      tableXml({
        tblW: '<w:tblW w:w="0" w:type="auto"/>',
        grid: '<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>',
      })
    );
    expect(total(widths)).toBeCloseTo(CONTENT_WIDTH_PX, 0);
  });

  test('percent string form: tblW "100%" pct → full width', () => {
    const widths = resolveWidths(
      tableXml({
        tblW: '<w:tblW w:w="100%" w:type="pct"/>',
        grid: '<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>',
      })
    );
    expect(total(widths)).toBeCloseTo(CONTENT_WIDTH_PX, 0);
  });

  test('percent string form on cells: tcW "50%" each → full width, even columns', () => {
    const widths = resolveWidths(
      tableXml({
        tblW: '<w:tblW w:w="100%" w:type="pct"/>',
        tcw: '<w:tcW w:w="50%" w:type="pct"/>',
      })
    );
    expect(total(widths)).toBeCloseTo(CONTENT_WIDTH_PX, 0);
    expect(widths[0]).toBeCloseTo(widths[1], 0);
  });

  test('fixed layout with a genuinely narrow grid stays narrow (Word-faithful)', () => {
    const widths = resolveWidths(
      tableXml({
        tblW: '<w:tblW w:w="0" w:type="auto"/>',
        grid: '<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>',
        layout: '<w:tblLayout w:type="fixed"/>',
      })
    );
    expect(total(widths)).toBeLessThan(20);
  });

  test('control: proper dxa grid unchanged', () => {
    const widths = resolveWidths(
      tableXml({
        tblW: '<w:tblW w:w="9360" w:type="dxa"/>',
        grid: '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>',
      })
    );
    expect(widths[0]).toBeCloseTo(312, 0);
    expect(widths[1]).toBeCloseTo(312, 0);
  });
});
