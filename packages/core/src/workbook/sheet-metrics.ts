import { attr, children, descendants, parseXmlTree } from '../xml/xml-tree.js';

import type { PackageReader } from '../package/package-reader.js';

export interface SheetMetrics {
  readonly defaultColumnWidthPx: number;
  readonly defaultRowHeightPx: number;
  readonly columnWidthsPx: ReadonlyMap<number, number>;
  readonly rowHeightsPx: ReadonlyMap<number, number>;
  readonly hiddenColumns: ReadonlySet<number>;
  readonly hiddenRows: ReadonlySet<number>;
}

const defaultExcelColumnWidth = 8.43;
const defaultRowHeightPt = 15;
const pxPerPoint = 96 / 72;

export async function parseSheetMetrics(
  packageReader: PackageReader,
  sheetPart: string
): Promise<SheetMetrics> {
  if (!(await packageReader.hasPart(sheetPart))) {
    return createDefaultSheetMetrics();
  }

  const xml = (await packageReader.readPart(sheetPart)).toString('utf8');
  const sheet = parseXmlTree(xml, sheetPart);
  const sheetFormat = descendants(sheet, 'sheetFormatPr')[0];
  const defaultColumnWidthPx = excelColumnWidthToPx(
    parseNumberAttr(sheetFormat, 'defaultColWidth') ?? defaultExcelColumnWidth
  );
  const defaultRowHeightPx = pointsToPx(
    parseNumberAttr(sheetFormat, 'defaultRowHeight') ?? defaultRowHeightPt
  );
  const columnWidthsPx = new Map<number, number>();
  const hiddenColumns = new Set<number>();

  for (const col of descendants(sheet, 'col')) {
    const min = parseNumberAttr(col, 'min');
    const max = parseNumberAttr(col, 'max');
    if (!min || !max) {
      continue;
    }

    const width = parseNumberAttr(col, 'width');
    const hidden = attr(col, 'hidden') === '1' || attr(col, 'hidden') === 'true';
    for (let column = min; column <= max; column += 1) {
      const zeroBasedColumn = column - 1;
      if (hidden) {
        hiddenColumns.add(zeroBasedColumn);
        columnWidthsPx.set(zeroBasedColumn, 0);
      } else if (width !== undefined) {
        columnWidthsPx.set(zeroBasedColumn, excelColumnWidthToPx(width));
      }
    }
  }

  const rowHeightsPx = new Map<number, number>();
  const hiddenRows = new Set<number>();
  for (const row of children(sheet, 'sheetData').flatMap((sheetData) => children(sheetData, 'row'))) {
    const rowIndex = parseNumberAttr(row, 'r');
    if (!rowIndex) {
      continue;
    }

    const zeroBasedRow = rowIndex - 1;
    if (attr(row, 'hidden') === '1' || attr(row, 'hidden') === 'true') {
      hiddenRows.add(zeroBasedRow);
      rowHeightsPx.set(zeroBasedRow, 0);
      continue;
    }

    const height = parseNumberAttr(row, 'ht');
    if (height !== undefined) {
      rowHeightsPx.set(zeroBasedRow, pointsToPx(height));
    }
  }

  return {
    defaultColumnWidthPx,
    defaultRowHeightPx,
    columnWidthsPx,
    rowHeightsPx,
    hiddenColumns,
    hiddenRows
  };
}

export function createDefaultSheetMetrics(): SheetMetrics {
  return {
    defaultColumnWidthPx: excelColumnWidthToPx(defaultExcelColumnWidth),
    defaultRowHeightPx: pointsToPx(defaultRowHeightPt),
    columnWidthsPx: new Map<number, number>(),
    rowHeightsPx: new Map<number, number>(),
    hiddenColumns: new Set<number>(),
    hiddenRows: new Set<number>()
  };
}

export function columnWidthPx(metrics: SheetMetrics, zeroBasedColumn: number): number {
  return metrics.columnWidthsPx.get(zeroBasedColumn) ?? metrics.defaultColumnWidthPx;
}

export function rowHeightPx(metrics: SheetMetrics, zeroBasedRow: number): number {
  return metrics.rowHeightsPx.get(zeroBasedRow) ?? metrics.defaultRowHeightPx;
}

export function excelColumnWidthToPx(width: number): number {
  if (!Number.isFinite(width) || width <= 0) {
    return 0;
  }
  return Math.floor(width * 7 + 5);
}

function pointsToPx(points: number): number {
  return points * pxPerPoint;
}

function parseNumberAttr(
  node: { readonly attributes: ReadonlyMap<string, string> } | undefined,
  name: string
): number | undefined {
  const value = node?.attributes.get(name);
  if (!value) {
    return undefined;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}
