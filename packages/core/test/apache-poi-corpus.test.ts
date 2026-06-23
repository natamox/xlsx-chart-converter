import fs from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';
import { openWorkbook } from '../src/index.js';

interface ExpectedWorkbook {
  readonly charts: number;
  readonly types: readonly string[];
  readonly minPoints: number;
}

const corpusDir = path.resolve(import.meta.dirname, '../../../fixtures/workbooks/apache-poi');
const expected = new Map<string, ExpectedWorkbook>([
  ['123233_charts.xlsx', { charts: 4, types: ['column', 'pie', 'column', 'line'], minPoints: 34 }],
  ['60255_extra_drawingparts.xlsx', { charts: 1, types: ['column'], minPoints: 8 }],
  ['chartTitle_noTitle.xlsx', { charts: 1, types: ['scatter'], minPoints: 4 }],
  ['chartTitle_withTitle.xlsx', { charts: 1, types: ['scatter'], minPoints: 4 }],
  ['chartTitle_withTitleFormula.xlsx', { charts: 1, types: ['scatter'], minPoints: 4 }],
  ['chart_sheet.xlsx', { charts: 1, types: ['column'], minPoints: 8 }],
  ['SimpleScatterChart.xlsx', { charts: 2, types: ['scatter', 'scatter'], minPoints: 4 }],
  ['WithChart.xlsx', { charts: 1, types: ['line'], minPoints: 12 }],
  ['WithChartSheet.xlsx', { charts: 1, types: ['column'], minPoints: 6 }],
  ['WithDrawing.xlsx', { charts: 0, types: [], minPoints: 0 }],
  ['WithThreeCharts.xlsx', { charts: 3, types: ['line', 'pie', 'area'], minPoints: 30 }],
  ['WithTwoCharts.xlsx', { charts: 2, types: ['line', 'area'], minPoints: 24 }]
]);

describe('Apache POI workbook corpus', () => {
  for (const [fileName, workbookExpectation] of expected) {
    it(`parses ${fileName}`, async () => {
      const workbook = await openWorkbook({ path: path.join(corpusDir, fileName) });

      const charts = await workbook.listCharts();
      const models = [];
      for (const chart of charts) {
        models.push(await workbook.getChartModel(chart.id));
      }

      await workbook.close();

      expect(charts).toHaveLength(workbookExpectation.charts);
      expect(charts.map((chart) => chart.chartTypes[0])).toEqual(workbookExpectation.types);
      expect(models.every((model) => model.plotArea?.chartGroups.length === model.chartTypes.length)).toBe(true);
      expect(models.every((model) => Array.isArray(model.axes))).toBe(true);
      expect(models.reduce((sum, model) =>
        sum + model.series.reduce((seriesSum, series) => seriesSum + series.points.length, 0), 0
      )).toBeGreaterThanOrEqual(workbookExpectation.minPoints);
    });
  }

  it('tracks every non-temporary workbook in the corpus', () => {
    const files = fs.readdirSync(corpusDir)
      .filter((fileName) => fileName.endsWith('.xlsx') && !fileName.startsWith('~$'))
      .sort();

    expect(files).toEqual([...expected.keys()].sort());
  });
});
