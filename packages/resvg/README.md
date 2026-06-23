# @natamox/excel-chart-resvg

resvg adapter for rasterizing processed excel-chart SVG output to PNG.

This package depends on `@natamox/excel-chart-core` for PNG renderer interfaces and on `@resvg/resvg-js` for rasterization.

## Usage

```ts
import { ResvgPngRenderer } from '@natamox/excel-chart-resvg';

const renderer = new ResvgPngRenderer();
```

## Status

The package is currently an M0 scaffold. The adapter class and package boundary exist; actual PNG rasterization options, font loading, and background handling are still being implemented.

## Runtime

Node.js >= 22.
