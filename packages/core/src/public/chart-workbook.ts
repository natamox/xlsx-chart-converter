import { parseClassicChartPart } from '../chart/classic/chart-space-parser.js';
import { locateChartsInDrawing } from '../chart/locator/chart-locator.js';
import { createExcelJsDataProvider } from '../data/exceljs-data-provider.js';
import { UnsupportedOperationError, XlsxChartError } from '../diagnostics/errors.js';
import { yauzlPackageReaderFactory } from '../package/yauzl-package-reader.js';
import { parseWorkbookIndex } from '../workbook/workbook-parser.js';
import { createWorksheetDataProvider } from '../workbook/worksheet-data-provider.js';
import type {
  ChartDescriptor,
  ChartEngineOptions,
  ChartId,
  ChartModel,
  DataMode,
  ExcelJsWorkbookLike,
  GetChartModelOptions,
  Diagnostic,
  PngRenderer,
  RenderContext,
  RenderOptions,
  RenderResult,
  WorkbookHandle,
  WorkbookSource
} from './types.js';
import type { PackageReader } from '../package/package-reader.js';
import type { WorkbookIndex } from '../workbook/workbook-index.js';
import type { WorksheetDataProvider } from '../workbook/worksheet-data-provider.js';
import type { ChartDataProvider } from '../data/data-resolver.js';

export class ChartWorkbook implements WorkbookHandle {
  static async open(
    source: WorkbookSource,
    options: ChartEngineOptions = {}
  ): Promise<ChartWorkbook> {
    const normalizedSource = normalizeSource(source);
    const packageReader = await yauzlPackageReaderFactory.open(normalizedSource);
    try {
      const workbookIndex = await parseWorkbookIndex(packageReader);
      return new ChartWorkbook(packageReader, workbookIndex, options, sourceWorkbook(source));
    } catch (error) {
      await packageReader.close();
      throw error;
    }
  }

  private closed = false;
  private descriptors?: readonly ChartDescriptor[];
  private readonly modelCache = new Map<string, ChartModel>();
  private readonly fallbackDataProvider: WorksheetDataProvider;
  private readonly exceljsDataProvider: ChartDataProvider | undefined;

  private constructor(
    private readonly packageReader: PackageReader,
    private readonly workbookIndex: WorkbookIndex,
    private readonly options: ChartEngineOptions,
    exceljsWorkbook: ExcelJsWorkbookLike | undefined
  ) {
    this.fallbackDataProvider = createWorksheetDataProvider(packageReader, workbookIndex);
    this.exceljsDataProvider = exceljsWorkbook ? createExcelJsDataProvider(exceljsWorkbook) : undefined;
  }

  async listCharts(): Promise<ChartDescriptor[]> {
    this.assertOpen();
    if (this.descriptors) {
      return [...this.descriptors];
    }

    const descriptors: ChartDescriptor[] = [];
    for (const sheet of this.workbookIndex.sheets) {
      descriptors.push(...await locateChartsInDrawing({
        packageReader: this.packageReader,
        sheet
      }));
    }

    this.descriptors = descriptors.map((descriptor, index) => ({
      ...descriptor,
      id: descriptor.id || `chart-${index + 1}`
    }));
    return [...this.descriptors];
  }

  async getChartModel(chartId: ChartId, options: GetChartModelOptions = {}): Promise<ChartModel> {
    this.assertOpen();
    const dataMode = options.dataMode ?? this.options.defaultDataMode ?? 'chart-cache-first';
    const cacheKey = `${chartId}:${dataMode}`;
    const cached = this.modelCache.get(cacheKey);
    if (cached && isModelCacheable(dataMode)) {
      return cached;
    }

    const descriptor = await this.requireDescriptor(chartId);
    if (!descriptor.supported) {
      throw new UnsupportedOperationError(
        'ERR_UNSUPPORTED_CHART',
        `Chart "${chartId}" is not supported by this version.`
      );
    }

    const parseOptions = {
      packageReader: this.packageReader,
      descriptor,
      dataMode,
      workbookIndex: this.workbookIndex,
      fallbackDataProvider: this.fallbackDataProvider,
      ...(this.exceljsDataProvider ? { exceljsDataProvider: this.exceljsDataProvider } : {})
    };
    const model = await parseClassicChartPart(parseOptions);
    if (isModelCacheable(dataMode)) {
      this.modelCache.set(cacheKey, model);
    }
    return model;
  }

  async render(chartId: ChartId, options: RenderOptions): Promise<RenderResult> {
    this.assertOpen();
    const model = await this.getChartModel(
      chartId,
      options.dataMode ? { dataMode: options.dataMode } : {}
    );
    const context = {
      width: options.width ?? model.width,
      height: options.height ?? model.height,
      scale: options.scale ?? 1,
      ...(options.background ? { background: options.background } : {}),
      ...(model.style?.fonts
        ? { fonts: { families: [model.style.fonts.minorLatin, model.style.fonts.majorLatin].filter(isString) } }
        : {})
    };
    const svg = this.options.renderer
      ? await this.options.renderer.render(model, context)
      : renderFallbackSvg(model, context.width, context.height);

    if (options.format === 'png') {
      if (!this.options.pngRenderer) {
        throw new UnsupportedOperationError(
          'ERR_PNG_RENDERER_REQUIRED',
          'PNG rendering requires a pngRenderer option.'
        );
      }
      const pngResult = await renderPng(this.options.pngRenderer, svg, context);
      return {
        chartId,
        ...(model.name === undefined ? {} : { chartName: model.name }),
        format: 'png',
        mediaType: 'image/png',
        data: pngResult.data,
        width: Math.round(context.width * context.scale),
        height: Math.round(context.height * context.scale),
        scale: context.scale,
        diagnostics: [...model.diagnostics, ...pngResult.diagnostics]
      };
    }

    return {
      chartId,
      ...(model.name === undefined ? {} : { chartName: model.name }),
      format: 'svg',
      mediaType: 'image/svg+xml',
      data: svg,
      width: context.width,
      height: context.height,
      scale: context.scale,
      diagnostics: model.diagnostics
    };
  }

  async close(): Promise<void> {
    this.closed = true;
    await this.packageReader.close();
  }

  private assertOpen(): void {
    if (this.closed) {
      throw new UnsupportedOperationError('ERR_WORKBOOK_CLOSED', 'Workbook handle is closed.');
    }
  }

  private async requireDescriptor(chartId: ChartId): Promise<ChartDescriptor> {
    const descriptors = await this.listCharts();
    const descriptor = descriptors.find((item) => item.id === chartId || item.name === chartId);
    if (!descriptor) {
      throw new XlsxChartError('ERR_CHART_NOT_FOUND', `Chart "${chartId}" was not found.`);
    }
    return descriptor;
  }
}

function isString(value: string | undefined): value is string {
  return typeof value === 'string' && value.length > 0;
}

async function renderPng(
  renderer: PngRenderer,
  svg: string,
  context: RenderContext
): Promise<{ data: Buffer; diagnostics: readonly Diagnostic[] }> {
  return renderer.renderWithDiagnostics
    ? renderer.renderWithDiagnostics(svg, context)
    : { data: await renderer.render(svg, context), diagnostics: [] };
}

function normalizeSource(source: WorkbookSource): Buffer | Uint8Array | { path: string } {
  if (Buffer.isBuffer(source) || source instanceof Uint8Array || 'path' in source) {
    return source;
  }
  return source.buffer;
}

function sourceWorkbook(source: WorkbookSource): ExcelJsWorkbookLike | undefined {
  return Buffer.isBuffer(source) || source instanceof Uint8Array ? undefined : source.workbook;
}

function isModelCacheable(dataMode: DataMode): boolean {
  return dataMode !== 'exceljs-first' && dataMode !== 'exceljs-only';
}

function renderFallbackSvg(model: ChartModel, width: number, height: number): string {
  const title = escapeXml(model.title ?? model.name ?? model.id);
  const values = model.series.flatMap((series) => series.points.map((point) => point.value ?? 0));
  const max = Math.max(1, ...values);
  const bars = values.slice(0, 20).map((value, index) => {
    const barWidth = Math.max(1, (width - 80) / Math.max(1, Math.min(values.length, 20)));
    const barHeight = Math.max(1, (height - 90) * (value / max));
    const x = 40 + index * barWidth;
    const y = height - 40 - barHeight;
    return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${Math.max(1, barWidth - 4).toFixed(1)}" height="${barHeight.toFixed(1)}" fill="#3b82f6"/>`;
  }).join('');

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" role="img" aria-label="${title}"><rect width="100%" height="100%" fill="#fff"/><text x="20" y="28" font-family="Arial, sans-serif" font-size="18" fill="#111">${title}</text>${bars}</svg>`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}
