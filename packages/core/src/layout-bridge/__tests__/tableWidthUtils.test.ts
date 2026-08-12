import { describe, test, expect } from 'bun:test';
import {
  resolveTableWidthPx,
  normalizeTableColumnWidths,
  resolveTableColumnWidths,
} from '../tableWidthUtils';
import type { TableBlock } from '../../layout-engine';

describe('resolveTableWidthPx', () => {
  test('dxa: twips converted to pixels', () => {
    // 1440 twips = 1 inch = 96 px
    expect(resolveTableWidthPx(1440, 'dxa', 600)).toBeCloseTo(96, 1);
  });

  test('pct: 50ths of a percent (5000 = 100%) per ECMA-376 §17.18.111', () => {
    expect(resolveTableWidthPx(2500, 'pct', 600)).toBe(300);
    expect(resolveTableWidthPx(5000, 'pct', 600)).toBe(600);
    // Small spec values must NOT be coerced to plain percent — `1` means 0.02%.
    expect(resolveTableWidthPx(1, 'pct', 5000)).toBeCloseTo(1, 5);
  });

  test('zero / negative / undefined width returns undefined', () => {
    expect(resolveTableWidthPx(0, 'dxa', 600)).toBeUndefined();
    expect(resolveTableWidthPx(-10, 'dxa', 600)).toBeUndefined();
    expect(resolveTableWidthPx(undefined, 'dxa', 600)).toBeUndefined();
  });

  test('unrecognized widthType returns undefined', () => {
    expect(resolveTableWidthPx(1440, 'nil', 600)).toBeUndefined();
  });
});

describe('normalizeTableColumnWidths', () => {
  test('empty array returns evenly-split target width', () => {
    expect(normalizeTableColumnWidths([], 3, 300)).toEqual([100, 100, 100]);
  });

  test('missing trailing columns inherit average of existing positives', () => {
    expect(normalizeTableColumnWidths([100, 200], 4, 1000)).toEqual([100, 200, 150, 150]);
  });

  test('zero/negative widths split the leftover target evenly', () => {
    const out = normalizeTableColumnWidths([100, 0, 100, -5], 4, 400);
    expect(out[0]).toBe(100);
    expect(out[2]).toBe(100);
    expect(out[1]).toBeCloseTo(100, 5);
    expect(out[3]).toBeCloseTo(100, 5);
  });

  test('all zero returns even split of target', () => {
    expect(normalizeTableColumnWidths([0, 0, 0], 3, 300)).toEqual([100, 100, 100]);
  });
});

describe('resolveTableColumnWidths — degenerate autofit grids', () => {
  function makeTable(overrides: Partial<TableBlock>, cols = 3): TableBlock {
    const cells = Array.from({ length: cols }, (_, i) => ({ id: i, blocks: [] }));
    return {
      kind: 'table',
      id: 0,
      rows: [{ id: 100, cells }],
      ...overrides,
    };
  }

  test('autofit + no width + placeholder grid (docx npm gridCol w=100) → even split', () => {
    // 100 twips ≈ 6.7px per column — narrower than the default cell margins.
    const table = makeTable({ columnWidths: [6.7, 6.7, 6.7], width: 0, widthType: 'auto' });
    const widths = resolveTableColumnWidths(table, 624);
    expect(widths).toHaveLength(3);
    for (const w of widths) expect(w).toBeCloseTo(208, 0);
  });

  test('fixed layout keeps a narrow grid literally (Word honors it)', () => {
    const table = makeTable({
      columnWidths: [6.7, 6.7, 6.7],
      width: 0,
      widthType: 'auto',
      tableLayout: 'fixed',
    });
    const widths = resolveTableColumnWidths(table, 624);
    for (const w of widths) expect(w).toBeCloseTo(6.7, 1);
  });

  test('a single narrow spacer column among wide ones is preserved', () => {
    const table = makeTable({ columnWidths: [7, 300, 317], width: 0, widthType: 'auto' });
    expect(resolveTableColumnWidths(table, 624)).toEqual([7, 300, 317]);
  });

  test('explicit table width rescales a narrow grid instead of discarding it', () => {
    // 9360 twips = 624px target; ratios (all equal) spread across it.
    const table = makeTable({ columnWidths: [6.7, 6.7, 6.7], width: 9360, widthType: 'dxa' });
    const widths = resolveTableColumnWidths(table, 624);
    for (const w of widths) expect(w).toBeCloseTo(208, 0);
  });
});
