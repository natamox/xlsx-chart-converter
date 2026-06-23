import { addressKey, parseA1Reference } from './a1-reference-parser.js';

import type { DataProviderContext, DataProviderResult, CellValue } from '../workbook/worksheet-data-provider.js';
import type { Diagnostic, ExcelJsWorkbookLike } from '../public/types.js';

export interface ExcelJsDataProvider {
  readonly getRangeValues: (formula: string, context?: DataProviderContext) => Promise<DataProviderResult>;
  readonly getSingleValue: (formula: string, context?: DataProviderContext) => Promise<CellValue | undefined>;
}

export function createExcelJsDataProvider(workbook: ExcelJsWorkbookLike): ExcelJsDataProvider {
  function getRangeValues(
    formula: string,
    context: DataProviderContext = {}
  ): Promise<DataProviderResult> {
    const reference = parseA1Reference(formula, context.sheetName);
    if (!reference?.sheetName) {
      return Promise.resolve(unresolved(formula, 'Formula is not a supported simple A1 reference.'));
    }

    const sheet = workbook.getWorksheet(reference.sheetName);
    if (!sheet) {
      return Promise.resolve(unresolved(formula, 'Referenced worksheet was not found in the ExcelJS workbook.'));
    }

    const values: CellValue[] = [];
    const diagnostics: Diagnostic[] = [];
    for (let row = reference.start.row; row <= reference.end.row; row += 1) {
      if (!context.includeHidden && sheet.getRow(row).hidden) {
        continue;
      }
      for (let col = reference.start.col; col <= reference.end.col; col += 1) {
        if (!context.includeHidden && sheet.getColumn(col).hidden) {
          continue;
        }

        const value = normalizeExcelJsValue(sheet.getCell(addressKey({ row, col })).value);
        if (value.kind === 'formula') {
          diagnostics.push({
            code: 'STALE_FORMULA_RESULT_POSSIBLE',
            severity: 'warning',
            message: 'ExcelJS exposes cached formula results but does not recalculate formulas.',
            details: { formula, cell: addressKey({ row, col }) }
          });
        }
        if (value.value !== undefined) {
          values.push(value.value);
        }
      }
    }

    return Promise.resolve({ values, diagnostics, resolved: true });
  }

  async function getSingleValue(
    formula: string,
    context: DataProviderContext = {}
  ): Promise<CellValue | undefined> {
    return (await getRangeValues(formula, context)).values[0];
  }

  return { getRangeValues, getSingleValue };
}

function normalizeExcelJsValue(value: unknown): { value?: CellValue; kind?: 'formula' } {
  if (value === null || value === undefined) {
    return {};
  }
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return { value };
  }
  if (value instanceof Date) {
    return { value: value.getTime() };
  }
  if (typeof value === 'object' && 'result' in value) {
    const result = normalizeExcelJsValue((value as { result?: unknown }).result).value;
    return result === undefined ? { kind: 'formula' } : { value: result, kind: 'formula' };
  }
  if (typeof value === 'object' && 'richText' in value) {
    const text = (value as { richText?: readonly { text?: unknown }[] }).richText
      ?.map((item) => primitiveToString(item.text))
      .join('');
    return text ? { value: text } : {};
  }
  if (typeof value === 'object' && 'text' in value) {
    return { value: primitiveToString((value as { text?: unknown }).text) };
  }
  return {};
}

function primitiveToString(value: unknown): string {
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
    ? String(value)
    : '';
}

function unresolved(formula: string, message: string): DataProviderResult {
  return {
    values: [],
    diagnostics: [{
      code: 'DATA_REFERENCE_UNRESOLVED',
      severity: 'warning',
      message,
      details: { formula }
    }],
    resolved: false
  };
}
