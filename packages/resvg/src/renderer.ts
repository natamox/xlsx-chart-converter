import { UnsupportedOperationError } from '@natamox/excel-chart-core';

import type { PngRenderer, RenderContext } from '@natamox/excel-chart-core';

export interface ResvgPngRendererOptions {
  readonly defaultBackground?: string;
}

export class ResvgPngRenderer implements PngRenderer {
  constructor(private readonly options: ResvgPngRendererOptions = {}) {
    void this.options;
  }

  render(svg: string, context: RenderContext): Promise<Buffer> {
    void svg;
    void context;
    return Promise.reject(new UnsupportedOperationError(
      'ERR_NOT_IMPLEMENTED',
      'resvg PNG rendering is planned for M5 and is not implemented yet.'
    ));
  }
}
