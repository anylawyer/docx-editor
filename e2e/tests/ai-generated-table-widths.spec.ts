/**
 * Tables from AI document generators must render readable, not as a sliver.
 *
 * The fixture (scripts/create-ai-generated-table-fixture.mjs) carries the two
 * shapes such generators emit and Word autofits on open:
 *  1. `docx` npm default — `w:tblW w:type="auto" w:w="0"` plus a `w:tblGrid`
 *     of 100-twip placeholder columns;
 *  2. `%`-suffixed ST_MeasurementOrPercent widths — `w:tblW w:w="100%"` and
 *     per-cell `w:tcW w:w="50%"`.
 * Both used to paint ~13px-wide tables with text wrapped one character per
 * line. The painter is shared core, so React coverage here covers Vue too.
 */
import { test, expect } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';

const FIXTURE = 'fixtures/ai-generated-table-widths.docx';

test.describe('AI-generated table widths', () => {
  test('autofit placeholder grid and percent tables render at usable width', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.goto();
    await editor.waitForReady();
    await page.locator('input[type="file"][accept=".docx"]').setInputFiles(`e2e/${FIXTURE}`);

    // Key on fixture-specific cell text so we know THIS document painted.
    await page.waitForFunction(() => {
      const cells = document.querySelectorAll('.layout-page-content .layout-table-cell');
      return [...cells].some((c) => (c.textContent ?? '').includes('Priority support'));
    });

    const { tableWidths, contentWidth } = await page.evaluate(() => {
      const cells = [
        ...document.querySelectorAll<HTMLElement>('.layout-page-content .layout-table-cell'),
      ];
      const tables = [...new Set(cells.map((c) => c.closest<HTMLElement>('[data-block-id]')))];
      return {
        tableWidths: tables
          .filter((t): t is HTMLElement => t !== null)
          .map((t) => t.getBoundingClientRect().width),
        contentWidth:
          document.querySelector<HTMLElement>('.layout-page-content')?.getBoundingClientRect()
            .width ?? 0,
      };
    });

    expect(tableWidths.length).toBeGreaterThanOrEqual(2);
    // Before the fix these painted at 20px / 12px on a 624px page.
    for (const width of tableWidths) {
      expect(width).toBeGreaterThan(contentWidth * 0.9);
    }
  });
});
