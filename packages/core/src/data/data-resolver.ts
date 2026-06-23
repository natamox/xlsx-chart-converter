import type { DataMode, Diagnostic } from '../public/types.js';
import type { DataProviderContext, DataProviderResult, CellValue } from '../workbook/worksheet-data-provider.js';

export interface ChartDataProvider {
  readonly getRangeValues: (formula: string, context?: DataProviderContext) => Promise<DataProviderResult>;
  readonly getSingleValue: (formula: string, context?: DataProviderContext) => Promise<CellValue | undefined>;
}

export interface ResolveChartDataOptions {
  readonly mode: DataMode;
  readonly formula?: string;
  readonly cacheValues: readonly CellValue[];
  readonly context?: DataProviderContext;
  readonly exceljsProvider?: ChartDataProvider;
  readonly fallbackProvider?: ChartDataProvider;
}

export interface ResolvedChartData {
  readonly values: readonly CellValue[];
  readonly source: 'cache' | 'exceljs' | 'fallback' | 'unresolved';
  readonly diagnostics: readonly Diagnostic[];
}

export async function resolveChartData(options: ResolveChartDataOptions): Promise<ResolvedChartData> {
  const cacheResult = fromCache(options.cacheValues);
  const diagnostics: Diagnostic[] = [];
  const attempts = resolutionOrder(options.mode);

  for (const attempt of attempts) {
    if (attempt === 'cache') {
      if (cacheResult) {
        return diagnostics.length > 0
          ? { ...cacheResult, diagnostics }
          : cacheResult;
      }
      continue;
    }

    if (!options.formula) {
      continue;
    }

    const provider = attempt === 'exceljs' ? options.exceljsProvider : options.fallbackProvider;
    if (!provider) {
      continue;
    }

    const result = await provider.getRangeValues(options.formula, options.context);
    diagnostics.push(...result.diagnostics);
    if (result.resolved && result.values.length > 0) {
      return {
        values: result.values,
        source: attempt,
        diagnostics
      };
    }
  }

  return {
    values: [],
    source: 'unresolved',
    diagnostics: diagnostics.length > 0 ? diagnostics : [{
      code: 'DATA_REFERENCE_UNRESOLVED',
      severity: options.mode === 'exceljs-only' ? 'error' : 'warning',
      message: 'Chart data could not be resolved from the configured data mode.',
      ...(options.formula ? { details: { formula: options.formula, mode: options.mode } } : {})
    }]
  };
}

function fromCache(cacheValues: readonly CellValue[]): ResolvedChartData | undefined {
  return cacheValues.length > 0
    ? { values: cacheValues, source: 'cache', diagnostics: [] }
    : undefined;
}

function resolutionOrder(mode: DataMode): readonly ('cache' | 'exceljs' | 'fallback')[] {
  if (mode === 'exceljs-first') {
    return ['exceljs', 'cache', 'fallback'];
  }
  if (mode === 'cache-only') {
    return ['cache'];
  }
  if (mode === 'exceljs-only') {
    return ['exceljs', 'fallback'];
  }
  return ['cache', 'exceljs', 'fallback'];
}
