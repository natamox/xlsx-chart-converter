# @natamox/excel-chart-core

Core package for Excel OOXML chart discovery, parsing, diagnostics, and public facade types.

This package intentionally does not depend on ECharts or resvg. It owns workbook/package boundaries, Chart IR types, diagnostics, and renderer interfaces so parsing users do not pay for rendering adapters.

## Usage

```ts
import { openWorkbook } from '@natamox/excel-chart-core';

const workbook = await openWorkbook(buffer);
const charts = await workbook.listCharts();
await workbook.close();
```

## Status

The package is currently an M0 scaffold. Workbook opening, public types, diagnostics, and package boundaries are present; chart discovery, parsing, and rendering integration are still being implemented.

## Runtime

Node.js >= 22.
