# @natamox/xlsx-chart-converter-resvg

resvg adapter for rasterizing processed xlsx-chart-converter SVG output to PNG.

This package depends on `@natamox/xlsx-chart-converter-core` for PNG renderer interfaces and on `@resvg/resvg-js` for rasterization.

## Usage

```ts
import { openWorkbook } from '@natamox/xlsx-chart-converter-core';
import { EChartsSvgRenderer } from '@natamox/xlsx-chart-converter-echarts';
import { ResvgPngRenderer } from '@natamox/xlsx-chart-converter-resvg';

const workbook = await openWorkbook(
  { path: 'report.xlsx' },
  {
    renderer: new EChartsSvgRenderer(),
    pngRenderer: new ResvgPngRenderer({ defaultBackground: '#fff' })
  }
);

const [chart] = await workbook.listCharts();
const png = await workbook.render(chart.id, {
  format: 'png',
  scale: 2
});

await workbook.close();
```

## Renderer Options

```ts
new ResvgPngRenderer({
  defaultBackground: '#fff'
});
```

Render context controls scale, background override, font families, and custom font files. Missing custom font files are returned as `FONT_FILE_NOT_FOUND` warnings from `renderWithDiagnostics()`.

## Implemented Capabilities

- PNG rasterization through `@resvg/resvg-js`.
- Scale-based output sizing.
- Optional background color.
- System font loading.
- Optional custom font file loading with diagnostics.

## Runtime

Node.js >= 22.
