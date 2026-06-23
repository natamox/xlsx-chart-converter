import { describe, expect, it } from 'vitest';

import { resolveChartData } from '../src/data/data-resolver.js';

describe('chart data resolver', () => {
  it('keeps provider diagnostics when exceljs-first falls back to chart cache', async () => {
    const result = await resolveChartData({
      mode: 'exceljs-first',
      formula: 'Sheet1!A1:A2',
      cacheValues: [1, 2],
      exceljsProvider: {
        getSingleValue: () => Promise.resolve(undefined),
        getRangeValues: () => Promise.resolve({
          values: [],
          resolved: false,
          diagnostics: [{
            code: 'DATA_REFERENCE_UNRESOLVED',
            severity: 'warning',
            message: 'ExcelJS data unavailable.'
          }]
        })
      }
    });

    expect(result.source).toBe('cache');
    expect(result.values).toEqual([1, 2]);
    expect(result.diagnostics.map((diagnostic) => diagnostic.code)).toEqual(['DATA_REFERENCE_UNRESOLVED']);
  });
});
