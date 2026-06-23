import { UnsupportedOperationError } from '@natamox/excel-chart-core';

import type { ChartModel, RenderContext, SvgRenderer } from '@natamox/excel-chart-core';

export interface EChartsRendererOptions {
  readonly rendererVersion?: string;
}

export class EChartsSvgRenderer implements SvgRenderer {
  constructor(private readonly options: EChartsRendererOptions = {}) {
    void this.options;
  }

  render(model: ChartModel, context: RenderContext): Promise<string> {
    void model;
    void context;
    return Promise.reject(new UnsupportedOperationError(
      'ERR_NOT_IMPLEMENTED',
      'ECharts SVG rendering is planned for M4 and is not implemented yet.'
    ));
  }
}
