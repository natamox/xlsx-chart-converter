import type { AnchorModel } from './anchor-model.js';
import type { SheetMetrics } from '../workbook/sheet-metrics.js';
import { columnWidthPx, createDefaultSheetMetrics, rowHeightPx } from '../workbook/sheet-metrics.js';

const emuPerInch = 914_400;
const defaultDpi = 96;
const fallbackWidth = 640;
const fallbackHeight = 360;

export function resolveAnchorSize(
  anchor: AnchorModel,
  metrics: SheetMetrics = createDefaultSheetMetrics()
): { width: number; height: number } {
  if (anchor.ext) {
    return clampSize({
      width: emuToPx(anchor.ext.cxEmu),
      height: emuToPx(anchor.ext.cyEmu)
    });
  }

  if (anchor.kind === 'twoCell' && anchor.from && anchor.to) {
    const x1 = markerX(anchor.from, metrics);
    const y1 = markerY(anchor.from, metrics);
    const x2 = markerX(anchor.to, metrics);
    const y2 = markerY(anchor.to, metrics);
    return clampSize({
      width: Math.abs(x2 - x1),
      height: Math.abs(y2 - y1)
    });
  }

  return {
    width: fallbackWidth,
    height: fallbackHeight
  };
}

export function emuToPx(emu: number, dpi = defaultDpi): number {
  return emu * dpi / emuPerInch;
}

function markerX(marker: { col: number; colOffEmu: number }, metrics: SheetMetrics): number {
  return sumColumns(metrics, marker.col) + emuToPx(marker.colOffEmu);
}

function markerY(marker: { row: number; rowOffEmu: number }, metrics: SheetMetrics): number {
  return sumRows(metrics, marker.row) + emuToPx(marker.rowOffEmu);
}

function sumColumns(metrics: SheetMetrics, endExclusive: number): number {
  let total = 0;
  for (let column = 0; column < endExclusive; column += 1) {
    total += columnWidthPx(metrics, column);
  }
  return total;
}

function sumRows(metrics: SheetMetrics, endExclusive: number): number {
  let total = 0;
  for (let row = 0; row < endExclusive; row += 1) {
    total += rowHeightPx(metrics, row);
  }
  return total;
}

function clampSize(size: { width: number; height: number }): { width: number; height: number } {
  const width = Number.isFinite(size.width) && size.width > 0 ? Math.round(size.width) : fallbackWidth;
  const height = Number.isFinite(size.height) && size.height > 0 ? Math.round(size.height) : fallbackHeight;
  return {
    width,
    height
  };
}
