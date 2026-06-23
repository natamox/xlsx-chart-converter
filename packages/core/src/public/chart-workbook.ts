import { UnsupportedOperationError } from '../diagnostics/errors.js';
import type {
  ChartDescriptor,
  ChartEngineOptions,
  ChartId,
  ChartModel,
  RenderOptions,
  RenderResult,
  WorkbookHandle,
  WorkbookSource
} from './types.js';

export class ChartWorkbook implements WorkbookHandle {
  static async open(
    source: WorkbookSource,
    options: ChartEngineOptions = {}
  ): Promise<ChartWorkbook> {
    return Promise.resolve(new ChartWorkbook(source, options));
  }

  private closed = false;

  private constructor(
    private readonly source: WorkbookSource,
    private readonly options: ChartEngineOptions
  ) {
    void this.source;
    void this.options;
  }

  async listCharts(): Promise<ChartDescriptor[]> {
    this.assertOpen();
    return Promise.resolve([]);
  }

  async getChartModel(chartId: ChartId): Promise<ChartModel> {
    this.assertOpen();
    return Promise.reject(new UnsupportedOperationError(
      'ERR_NOT_IMPLEMENTED',
      `Chart model parsing is not implemented yet for chart "${chartId}".`
    ));
  }

  async render(chartId: ChartId, options: RenderOptions): Promise<RenderResult> {
    this.assertOpen();
    void options;
    return Promise.reject(new UnsupportedOperationError(
      'ERR_NOT_IMPLEMENTED',
      `Rendering is not implemented yet for chart "${chartId}".`
    ));
  }

  async close(): Promise<void> {
    this.closed = true;
    return Promise.resolve();
  }

  private assertOpen(): void {
    if (this.closed) {
      throw new UnsupportedOperationError('ERR_WORKBOOK_CLOSED', 'Workbook handle is closed.');
    }
  }
}
