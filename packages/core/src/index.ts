export { createChartEngine, openWorkbook } from './public/create-chart-engine.js';
export { ChartWorkbook } from './public/chart-workbook.js';
export {
  UnsupportedOperationError,
  XlsxChartError,
  isXlsxChartError
} from './diagnostics/errors.js';
export type {
  ChartDescriptor,
  ChartEngine,
  ChartEngineOptions,
  ChartId,
  ChartModel,
  DataMode,
  Diagnostic,
  DiagnosticSeverity,
  RenderFormat,
  RenderContext,
  RenderOptions,
  RenderResult,
  PngRenderer,
  SvgRenderer,
  UnsupportedPolicy,
  WorkbookHandle,
  WorkbookSource
} from './public/types.js';
