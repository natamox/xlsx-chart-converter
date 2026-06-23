import { ChartWorkbook } from './chart-workbook.js';
import type { ChartEngine, ChartEngineOptions, WorkbookHandle, WorkbookSource } from './types.js';

export function createChartEngine(options: ChartEngineOptions = {}): ChartEngine {
  return {
    async open(source: WorkbookSource): Promise<WorkbookHandle> {
      return openWorkbook(source, options);
    }
  };
}

export async function openWorkbook(
  source: WorkbookSource,
  options: ChartEngineOptions = {}
): Promise<WorkbookHandle> {
  return ChartWorkbook.open(source, options);
}
