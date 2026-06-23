# @natamox/excel-chart-echarts

ECharts SSR adapter for rendering excel-chart IR to SVG.

This package depends on `@natamox/excel-chart-core` for Chart IR and renderer interfaces, and on `echarts` for server-side SVG rendering.

## Usage

```ts
import { EChartsSvgRenderer } from '@natamox/excel-chart-echarts';

const renderer = new EChartsSvgRenderer();
```

## Status

The package is currently an M0 scaffold. The adapter class and package boundary exist; actual ECharts option mapping and SVG rendering are still being implemented.

## Runtime

Node.js >= 22.
