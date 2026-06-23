export type WorkbookSource =
  | Buffer
  | Uint8Array
  | { path: string; workbook?: ExcelJsWorkbookLike }
  | { buffer: Buffer; workbook?: ExcelJsWorkbookLike };

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
  width?: number;
  height?: number;
  chartTypes: string[];
  supported: boolean;
  diagnostics: Diagnostic[];
}

export interface ChartModel {
  schemaVersion: 1;
  id: ChartId;
  name?: string;
  title?: string;
  width: number;
  height: number;
  chartTypes: string[];
  legend?: ChartLegend;
  plotArea?: ChartPlotArea;
  style?: ChartStyle;
  axes: ChartAxis[];
  series: ChartSeries[];
  diagnostics: Diagnostic[];
}

export type ChartAxisKind = 'category' | 'value' | 'date' | 'series';
export type ChartAxisPosition = 'left' | 'right' | 'top' | 'bottom';
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';

export interface ChartLegend {
  readonly position: 'left' | 'right' | 'top' | 'bottom' | 'corner' | 'unknown';
  readonly overlay: boolean;
  readonly layout?: ChartManualLayout;
  readonly textStyle?: ChartTextStyle;
}

export interface ChartPlotArea {
  readonly chartGroups: readonly ChartGroup[];
  readonly layout?: ChartManualLayout;
}

export type ChartLayoutTarget = 'inner' | 'outer';
export type ChartLayoutMode = 'edge' | 'factor';

export interface ChartManualLayout {
  readonly target?: ChartLayoutTarget;
  readonly xMode?: ChartLayoutMode;
  readonly yMode?: ChartLayoutMode;
  readonly widthMode?: ChartLayoutMode;
  readonly heightMode?: ChartLayoutMode;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
}

export interface ChartGroup {
  readonly type: string;
  readonly axisIds: readonly string[];
  readonly grouping?: ChartGrouping;
  readonly varyColors?: boolean;
  readonly dataLabels?: ChartDataLabels;
}

export interface ChartStyle {
  readonly chartArea?: ChartShapeStyle;
  readonly plotArea?: ChartShapeStyle;
  readonly series?: readonly ChartShapeStyle[];
  readonly seriesMarkers?: readonly (ChartMarkerStyle | undefined)[];
  readonly pointStyles?: readonly (readonly ChartPointStyle[] | undefined)[];
  readonly title?: ChartTextStyle;
  readonly legend?: ChartTextStyle;
  readonly axes?: readonly ChartAxisStyle[];
  readonly dataLabels?: ChartTextStyle;
  readonly fonts?: {
    readonly majorLatin?: string;
    readonly minorLatin?: string;
  };
}

export interface ChartShapeStyle {
  readonly fill?: ChartFillStyle;
  readonly line?: ChartLineStyle;
}

export type ChartFillKind = 'solid' | 'none' | 'gradient' | 'pattern' | 'picture';

export interface ChartFillStyle {
  readonly kind?: ChartFillKind;
  readonly color?: string;
  readonly transformedColor?: string;
  readonly alpha?: number;
}

export interface ChartLineStyle {
  readonly color?: string;
  readonly transformedColor?: string;
  readonly width?: number;
  readonly alpha?: number;
  readonly dash?: string;
  readonly noFill?: boolean;
}

export interface ChartTextStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string;
  readonly alpha?: number;
}

export interface ChartMarkerStyle {
  readonly symbol?: string;
  readonly size?: number;
  readonly fill?: ChartFillStyle;
  readonly line?: ChartLineStyle;
}

export interface ChartPointStyle {
  readonly index: number;
  readonly style?: ChartShapeStyle;
  readonly marker?: ChartMarkerStyle;
}

export interface ChartAxisStyle {
  readonly axisId?: string;
  readonly shape?: ChartShapeStyle;
  readonly text?: ChartTextStyle;
}

export interface ChartAxis {
  readonly id: string;
  readonly kind: ChartAxisKind;
  readonly position?: ChartAxisPosition;
  readonly title?: string;
  readonly style?: ChartAxisStyle;
  readonly crossesAxisId?: string;
  readonly numberFormat?: string;
  readonly scaling?: ChartAxisScaling;
}

export interface ChartAxisScaling {
  readonly orientation?: 'minMax' | 'maxMin';
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly minorUnit?: number;
  readonly logBase?: number;
}

export interface ChartDataLabels {
  readonly showValue?: boolean;
  readonly showCategoryName?: boolean;
  readonly showSeriesName?: boolean;
  readonly showLegendKey?: boolean;
  readonly showPercent?: boolean;
  readonly showBubbleSize?: boolean;
  readonly position?: string;
}

export interface ChartDataPoint {
  readonly category?: string;
  readonly x?: string | number;
  readonly y?: number;
  readonly value?: number;
}

export interface ChartSeries {
  readonly name: string;
  readonly chartType?: string;
  readonly axisIds?: readonly string[];
  readonly dataLabels?: ChartDataLabels;
  readonly points: readonly ChartDataPoint[];
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

export interface GetChartModelOptions {
  dataMode?: DataMode;
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
  renderWithDiagnostics?(svg: string, context: RenderContext): Promise<{
    readonly data: Buffer;
    readonly diagnostics: readonly Diagnostic[];
  }>;
}

export interface ChartEngineOptions {
  packageLimits?: Record<string, unknown>;
  defaultDataMode?: DataMode;
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
  getChartModel(chartId: ChartId, options?: GetChartModelOptions): Promise<ChartModel>;
  render(chartId: ChartId, options: RenderOptions): Promise<RenderResult>;
  close(): Promise<void>;
}

export interface ExcelJsWorkbookLike {
  getWorksheet(name: string): ExcelJsWorksheetLike | undefined;
}

export interface ExcelJsWorksheetLike {
  getCell(address: string): ExcelJsCellLike;
  getRow(row: number): { hidden?: boolean };
  getColumn(column: number): { hidden?: boolean };
}

export interface ExcelJsCellLike {
  value: unknown;
}
