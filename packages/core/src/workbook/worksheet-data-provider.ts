import { addressKey, parseA1Reference } from '../data/a1-reference-parser.js';
import { attr, child, children, descendants, parseXmlTree, textContent } from '../xml/xml-tree.js';

import type { PackageReader } from '../package/package-reader.js';
import type { Diagnostic } from '../public/types.js';
import type { WorkbookIndex, WorkbookSheetIndex } from './workbook-index.js';
import type { XmlNode } from '../xml/xml-tree.js';

export interface WorksheetDataProvider {
  readonly getRangeValues: (formula: string, context?: DataProviderContext) => Promise<DataProviderResult>;
  readonly getSingleValue: (formula: string, context?: DataProviderContext) => Promise<CellValue | undefined>;
}

export type CellValue = string | number | boolean;

export interface DataProviderContext {
  readonly sheetName?: string;
  readonly includeHidden?: boolean;
}

export interface DataProviderResult {
  readonly values: readonly CellValue[];
  readonly diagnostics: readonly Diagnostic[];
  readonly resolved: boolean;
}

export function createWorksheetDataProvider(
  packageReader: PackageReader,
  workbookIndex: WorkbookIndex
): WorksheetDataProvider {
  const sharedStringsPromise = loadSharedStrings(packageReader, workbookIndex.sharedStringsPart);
  const sheetCache = new Map<string, Promise<LoadedSheetCells>>();

  async function getSheetCells(sheet: WorkbookSheetIndex): Promise<LoadedSheetCells> {
    let cached = sheetCache.get(sheet.partName);
    if (!cached) {
      cached = loadSheetCells(packageReader, sheet.partName, sharedStringsPromise);
      sheetCache.set(sheet.partName, cached);
    }
    return cached;
  }

  async function getRangeValues(
    formula: string,
    context: DataProviderContext = {}
  ): Promise<DataProviderResult> {
    const reference = parseA1Reference(formula, context.sheetName);
    if (!reference) {
      return {
        values: [],
        diagnostics: [unresolvedDiagnostic(formula, 'Formula is not a supported simple A1 reference.')],
        resolved: false
      };
    }

    const sheet = workbookIndex.sheets.find((item) => item.name === reference.sheetName);
    if (!sheet || sheet.kind !== 'worksheet') {
      return {
        values: [],
        diagnostics: [unresolvedDiagnostic(formula, 'Referenced worksheet was not found.')],
        resolved: false
      };
    }

    const loaded = await getSheetCells(sheet);
    const values: CellValue[] = [];
    for (let row = reference.start.row; row <= reference.end.row; row += 1) {
      if (!context.includeHidden && loaded.hiddenRows.has(row)) {
        continue;
      }
      for (let col = reference.start.col; col <= reference.end.col; col += 1) {
        if (!context.includeHidden && loaded.hiddenColumns.has(col)) {
          continue;
        }
        const value = loaded.cells.get(addressKey({ row, col }));
        if (value !== undefined) {
          values.push(value);
        }
      }
    }
    return {
      values,
      diagnostics: [...loaded.formulaCells].some((address) => isAddressInRange(address, reference.start, reference.end))
        ? [{
            code: 'STALE_FORMULA_RESULT_POSSIBLE',
            severity: 'warning',
            message: 'Worksheet formula results are cached values; ExcelJS does not recalculate formulas.',
            path: sheet.partName,
            details: { formula }
          }]
        : [],
      resolved: true
    };
  }

  async function getSingleValue(
    formula: string,
    context: DataProviderContext = {}
  ): Promise<CellValue | undefined> {
    return (await getRangeValues(formula, context)).values[0];
  }

  return { getRangeValues, getSingleValue };
}

async function loadSharedStrings(
  packageReader: PackageReader,
  sharedStringsPart: string | undefined
): Promise<readonly string[]> {
  if (!sharedStringsPart || !(await packageReader.hasPart(sharedStringsPart))) {
    return [];
  }

  const xml = (await packageReader.readPart(sharedStringsPart)).toString('utf8');
  const root = parseXmlTree(xml, sharedStringsPart);
  return children(root, 'si').map((item) => textContent(item) ?? '');
}

async function loadSheetCells(
  packageReader: PackageReader,
  sheetPart: string,
  sharedStringsPromise: Promise<readonly string[]>
): Promise<LoadedSheetCells> {
  const xml = (await packageReader.readPart(sheetPart)).toString('utf8');
  const root = parseXmlTree(xml, sheetPart);
  const sharedStrings = await sharedStringsPromise;
  const cells = new Map<string, CellValue>();
  const hiddenRows = new Set<number>();
  const hiddenColumns = new Set<number>();
  const formulaCells = new Set<string>();

  for (const colNode of descendants(root, 'col')) {
    if (attr(colNode, 'hidden') === '1' || attr(colNode, 'hidden') === 'true') {
      const min = Number(attr(colNode, 'min'));
      const max = Number(attr(colNode, 'max'));
      if (Number.isFinite(min) && Number.isFinite(max)) {
        for (let col = min; col <= max; col += 1) {
          hiddenColumns.add(col);
        }
      }
    }
  }

  for (const rowNode of descendants(root, 'row')) {
    const row = Number(attr(rowNode, 'r'));
    if (Number.isFinite(row) && (attr(rowNode, 'hidden') === '1' || attr(rowNode, 'hidden') === 'true')) {
      hiddenRows.add(row);
    }
  }

  for (const cellNode of descendants(root, 'c')) {
    const ref = attr(cellNode, 'r');
    if (!ref) {
      continue;
    }

    if (child(cellNode, 'f')) {
      formulaCells.add(ref.toUpperCase());
    }
    const value = parseCellValue(cellNode, sharedStrings);
    if (value !== undefined) {
      cells.set(ref.toUpperCase(), value);
    }
  }

  return { cells, hiddenRows, hiddenColumns, formulaCells };
}

function parseCellValue(cellNode: XmlNode, sharedStrings: readonly string[]): CellValue | undefined {
  const cellType = attr(cellNode, 't');
  if (cellType === 'inlineStr') {
    return textContent(child(cellNode, 'is'));
  }

  const raw = textContent(child(cellNode, 'v'));
  if (raw === undefined) {
    return undefined;
  }

  if (cellType === 's') {
    return sharedStrings[Number(raw)] ?? '';
  }

  if (cellType === 'b') {
    return raw === '1';
  }

  if (cellType === 'str') {
    return raw;
  }

  const numeric = Number(raw);
  return Number.isFinite(numeric) ? numeric : raw;
}

interface LoadedSheetCells {
  readonly cells: ReadonlyMap<string, CellValue>;
  readonly hiddenRows: ReadonlySet<number>;
  readonly hiddenColumns: ReadonlySet<number>;
  readonly formulaCells: ReadonlySet<string>;
}

function unresolvedDiagnostic(formula: string, message: string): Diagnostic {
  return {
    code: 'DATA_REFERENCE_UNRESOLVED',
    severity: 'warning',
    message,
    details: { formula }
  };
}

function isAddressInRange(address: string, start: { row: number; col: number }, end: { row: number; col: number }): boolean {
  const parsed = parseA1Reference(address, 'Sheet');
  if (!parsed) {
    return false;
  }
  return parsed.start.row >= start.row
    && parsed.start.row <= end.row
    && parsed.start.col >= start.col
    && parsed.start.col <= end.col;
}
