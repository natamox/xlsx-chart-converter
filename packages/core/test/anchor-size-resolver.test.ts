import { describe, expect, it } from 'vitest';

import { resolveAnchorSize } from '../src/drawing/anchor-size-resolver.js';

import type { AnchorModel } from '../src/drawing/anchor-model.js';
import type { SheetMetrics } from '../src/workbook/sheet-metrics.js';

describe('anchor size resolver', () => {
  it('uses sheet column widths and row heights for two-cell anchors', () => {
    const anchor: AnchorModel = {
      kind: 'twoCell',
      from: { col: 0, row: 0, colOffEmu: 0, rowOffEmu: 0 },
      to: { col: 3, row: 2, colOffEmu: 9_525, rowOffEmu: 19_050 }
    };
    const metrics: SheetMetrics = {
      defaultColumnWidthPx: 64,
      defaultRowHeightPx: 20,
      columnWidthsPx: new Map([
        [0, 100],
        [1, 0],
        [2, 50]
      ]),
      rowHeightsPx: new Map([
        [0, 12],
        [1, 24]
      ]),
      hiddenColumns: new Set([1]),
      hiddenRows: new Set()
    };

    expect(resolveAnchorSize(anchor, metrics)).toEqual({
      width: 151,
      height: 38
    });
  });

  it('falls back to explicit EMU extents for one-cell anchors', () => {
    expect(resolveAnchorSize({
      kind: 'oneCell',
      from: { col: 0, row: 0, colOffEmu: 0, rowOffEmu: 0 },
      ext: { cxEmu: 914_400, cyEmu: 457_200 }
    })).toEqual({
      width: 96,
      height: 48
    });
  });
});
