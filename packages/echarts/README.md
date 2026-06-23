# @natamox/xlsx-chart-converter-echarts

ECharts SSR adapter for rendering xlsx-chart-converter `ChartModel` IR to SVG.

This package depends on `@natamox/xlsx-chart-converter-core` for Chart IR and renderer interfaces, `@natamox/xlsx-chart-converter-svg` for post-processing, and `echarts` for server-side SVG rendering.

## Usage

```ts
import { openWorkbook } from '@natamox/xlsx-chart-converter-core';
import { EChartsSvgRenderer } from '@natamox/xlsx-chart-converter-echarts';

const workbook = await openWorkbook(
  { path: 'report.xlsx' },
  { renderer: new EChartsSvgRenderer() }
);

const [chart] = await workbook.listCharts();
const svg = await workbook.render(chart.id, {
  format: 'svg',
  width: chart.width,
  height: chart.height,
  background: '#fff'
});

await workbook.close();
```

## Renderer Options

```ts
new EChartsSvgRenderer({
  idPrefix: 'chart-',
  postProcess: true
});
```

- `idPrefix`: prefix for SVG IDs and `url(#id)` references.
- `postProcess`: enabled by default; sanitizes SVG, prefixes IDs, and adds accessibility metadata.

## Implemented Mapping

- Common cartesian, pie/doughnut, scatter, and mixed chart structures represented by the current IR.
- Category/value axes, secondary axes, axis scaling, number formats, and label formatters.
- Series data labels and blank point handling.
- F1 style fields where ECharts has equivalent options: fills, lines, dash, markers, per-point style, title/legend text style, and axis style.
- Deterministic SSR SVG with animation disabled.

## Limits

- ECharts is not a pixel-perfect Excel layout engine.
- Unsupported IR fields are intentionally ignored or degraded until F2/F3 fidelity work.
- Workbook-provided JavaScript formatters are never executed.

## Runtime

Node.js >= 22.
