# @natamox/excel-chart-core

Core package for Excel OOXML chart discovery, parsing, data resolution, theme/style resolution, diagnostics, and public facade types.

This package intentionally does not depend on ECharts or resvg. It owns workbook/package boundaries, Chart IR types, diagnostics, and renderer interfaces so parsing users do not pay for rendering adapters.

## Usage

```ts
import { openWorkbook } from '@natamox/excel-chart-core';

const workbook = await openWorkbook({ path: 'report.xlsx' });
try {
  const charts = await workbook.listCharts();
  const model = await workbook.getChartModel(charts[0].id, {
    dataMode: 'chart-cache-first'
  });
  const fallbackSvg = await workbook.render(model.id, { format: 'svg' });
} finally {
  await workbook.close();
}
```

For production SVG/PNG output, pass renderer adapters through `openWorkbook()` options.

## Public Surface

- `openWorkbook(source, options)`
- `createChartEngine(options)`
- `WorkbookHandle.listCharts()`
- `WorkbookHandle.getChartModel(chartId, options)`
- `WorkbookHandle.render(chartId, options)`
- `WorkbookHandle.close()`
- `ChartModel`, `ChartDescriptor`, renderer interfaces, diagnostics, data-mode, and style IR types

## Implemented Capabilities

- Safe access to the original OOXML ZIP package.
- Workbook, sheet, chartsheet, drawing, anchor, relationship, and chart discovery.
- Classic chart parsing into `ChartModel` schema version `1`.
- Chart cache, worksheet XML, and optional ExcelJS-like data provider resolution.
- A1 reference and range parsing, including quoted sheet names.
- Theme color/font/style resolution, including F1 style cascade fields.
- Structured diagnostics for unsupported or degraded behavior.
- Fallback SVG rendering when no renderer adapter is supplied.

## Limits

- `.xls`, `.xlsb`, encrypted workbooks, external OOXML relationships, VBA execution, and full formula calculation are not supported.
- PNG rendering requires a `PngRenderer` option.
- The fallback SVG is only a safety fallback; use `@natamox/excel-chart-echarts` for normal SVG rendering.

## Runtime

Node.js >= 22.
