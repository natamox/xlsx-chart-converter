import * as echarts from 'echarts';

import type { ChartModel, RenderContext, SvgRenderer } from '@natamox/excel-chart-core';
import { addSvgAccessibility, prefixSvgIds, sanitizeSvg } from '@natamox/excel-chart-svg';
import { buildEChartsOption } from './option-builder.js';

export interface EChartsRendererOptions {
  readonly rendererVersion?: string;
  readonly idPrefix?: string;
  readonly postProcess?: boolean;
}

export class EChartsSvgRenderer implements SvgRenderer {
  constructor(private readonly options: EChartsRendererOptions = {}) {
    void this.options;
  }

  render(model: ChartModel, context: RenderContext): Promise<string> {
    const chart = echarts.init(null, null, {
      renderer: 'svg',
      ssr: true,
      width: context.width,
      height: context.height
    });

    try {
      const builtOption = buildEChartsOption(model).option;
      const option = {
        animation: false,
        ...(context.fonts?.families?.[0] ? { textStyle: { fontFamily: context.fonts.families[0] } } : {}),
        ...builtOption,
        backgroundColor: context.background ?? builtOption.backgroundColor ?? 'transparent'
      };
      chart.setOption(option, true);
      const rawSvg = chart.renderToSVGString();
      return Promise.resolve(this.options.postProcess === false ? rawSvg : postProcessSvg(rawSvg, model, this.options.idPrefix));
    } finally {
      chart.dispose();
    }
  }
}

function postProcessSvg(svg: string, model: ChartModel, idPrefix: string | undefined): string {
  const prefix = idPrefix ?? `xc-${model.id.replaceAll(/[^a-zA-Z0-9_-]/g, '').slice(0, 24)}-`;
  return addSvgAccessibility(prefixSvgIds(sanitizeSvg(svg), prefix), model.title ?? model.name ?? model.id);
}
