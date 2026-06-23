import { describe, expect, it } from 'vitest';

import { addressKey, parseA1Reference } from '../src/data/a1-reference-parser.js';

describe('A1 reference parser', () => {
  it('parses quoted and non-ASCII sheet names', () => {
    expect(parseA1Reference("'销售 数据'!$B$2:$C$4")).toEqual({
      sheetName: '销售 数据',
      start: { col: 2, row: 2 },
      end: { col: 3, row: 4 }
    });
  });

  it('unescapes single quotes in sheet names', () => {
    expect(parseA1Reference("'Bob''s Sheet'!A1")).toMatchObject({
      sheetName: "Bob's Sheet",
      start: { col: 1, row: 1 },
      end: { col: 1, row: 1 }
    });
  });

  it('uses default sheet names for relative references and normalizes reversed ranges', () => {
    expect(parseA1Reference('$D$5:B2', 'Sheet1')).toEqual({
      sheetName: 'Sheet1',
      start: { col: 2, row: 2 },
      end: { col: 4, row: 5 }
    });
  });

  it('rejects references without a sheet context', () => {
    expect(parseA1Reference('A1')).toBeUndefined();
    expect(parseA1Reference('SUM(A1:A2)', 'Sheet1')).toBeUndefined();
  });

  it('formats address keys from numeric coordinates', () => {
    expect(addressKey({ col: 28, row: 10 })).toBe('AB10');
  });
});
