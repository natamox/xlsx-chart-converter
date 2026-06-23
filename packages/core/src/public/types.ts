export type WorkbookSource =
  | Buffer
  | Uint8Array
  | { path: string }
  | { buffer: Buffer; workbook?: unknown };

export type ChartId = string;

export type RenderFormat = 'svg' | 'png';

export type DataMode =
  | 'chart-cache-first'
  | 'exceljs-first'
  | 'cache-only'
  | 'exceljs-only';

export type UnsupportedPolicy = 'warn' | 'error' | 'degrade';

export type DiagnosticSeverity = 'info' | 'warning' | 'error';

export interface Diagnostic {
  code: string;
  severity: DiagnosticSeverity;
  message: string;
  path?: string;
  details?: Record<string, unknown>;
}

export interface ChartDescriptor {
  id: ChartId;
  name?: string;
  sheetName?: string;
  chartPart?: string;
  drawingPart?: string;
  chartTypes: string[];
  supported: boolean;
  diagnostics: Diagnostic[];
}

export interface ChartModel {
  schemaVersion: 1;
  id: ChartId;
  chartTypes: string[];
  diagnostics: Diagnostic[];
}

export interface RenderOptions {
  format: RenderFormat;
  width?: number;
  height?: number;
  scale?: number;
  background?: string;
  dataMode?: DataMode;
  unsupported?: UnsupportedPolicy;
}

export interface RenderResult {
  chartId?: ChartId;
  chartName?: string;
  format: RenderFormat;
  mediaType?: 'image/svg+xml' | 'image/png';
  data: string | Buffer;
  width?: number;
  height?: number;
  scale?: number;
  diagnostics: Diagnostic[];
}

export interface RenderContext {
  width: number;
  height: number;
  scale: number;
  background?: string;
  fonts?: {
    families?: readonly string[];
    fontFiles?: readonly string[];
  };
  signal?: AbortSignal;
}

export interface SvgRenderer {
  render(model: ChartModel, context: RenderContext): Promise<string>;
}

export interface PngRenderer {
  render(svg: string, context: RenderContext): Promise<Buffer>;
}

export interface ChartEngineOptions {
  packageLimits?: Record<string, unknown>;
  renderer?: SvgRenderer;
  pngRenderer?: PngRenderer;
  logger?: {
    debug?: (message: string, context?: Record<string, unknown>) => void;
    info?: (message: string, context?: Record<string, unknown>) => void;
    warn?: (message: string, context?: Record<string, unknown>) => void;
    error?: (message: string, context?: Record<string, unknown>) => void;
  };
}

export interface ChartEngine {
  open(source: WorkbookSource): Promise<WorkbookHandle>;
}

export interface WorkbookHandle {
  listCharts(): Promise<ChartDescriptor[]>;
  getChartModel(chartId: ChartId): Promise<ChartModel>;
  render(chartId: ChartId, options: RenderOptions): Promise<RenderResult>;
  close(): Promise<void>;
}
