# xlsx-chart-converter

Node.js workspace for discovering, parsing, inspecting, and rendering Excel OOXML charts from `.xlsx` and `.xlsm` workbooks.

The project is now past the initial scaffold. The core package can discover charts through workbook, sheet, drawing, and relationship parts; parse supported classic chart families into a deterministic Chart IR; resolve chart data and workbook/theme/style context; and render to SVG or PNG through adapter packages.

## Requirements

- Node.js >= 22
- pnpm 11.5.1

## Quick Start

```bash
pnpm install
pnpm run check
```

Export charts from the Apache POI fixture corpus for visual inspection:

```bash
make export-chart-corpus
```

The corpus preview is written to `output/chart-corpus-preview/index.html`. The `output/` directory is intentionally ignored by git.

## Common Commands

```bash
pnpm run lint
pnpm run test
pnpm run typecheck
pnpm run build
pnpm run check
make check
make export-chart-corpus
```

## Workspace Layout

```text
packages/
  core/      # OOXML package/workbook/drawing/chart parsing, Chart IR, diagnostics, facade
  echarts/   # ECharts server-side SVG renderer for Chart IR
  svg/       # SVG sanitizing, ID prefixing, accessibility post-processing
  resvg/     # resvg PNG rasterization adapter
  cli/       # xlsx-chart-converter command-line interface
fixtures/    # Apache POI workbook corpus and regression fixtures
scripts/     # local QA/export helpers
output/      # generated previews, ignored by git
```

Dependency direction:

```text
core      -> no renderer packages
echarts   -> core + svg + echarts
svg       -> core render types only where needed
resvg     -> core + @resvg/resvg-js
cli       -> core + echarts + svg + resvg
```

## Library Usage

```ts
import { openWorkbook } from '@natamox/xlsx-chart-converter-core';
import { EChartsSvgRenderer } from '@natamox/xlsx-chart-converter-echarts';
import { ResvgPngRenderer } from '@natamox/xlsx-chart-converter-resvg';

const workbook = await openWorkbook(
  { path: 'report.xlsx' },
  {
    renderer: new EChartsSvgRenderer(),
    pngRenderer: new ResvgPngRenderer()
  }
);

const charts = await workbook.listCharts();
const model = await workbook.getChartModel(charts[0].id);
const svg = await workbook.render(model.id, { format: 'svg' });
const png = await workbook.render(model.id, { format: 'png', scale: 2 });

await workbook.close();
```

Supported data modes:

- `chart-cache-first`
- `cache-only`
- `exceljs-first`
- `exceljs-only`

When rendering without an SVG renderer, `core` returns a simple fallback SVG. Production-quality SVG output should pass `EChartsSvgRenderer`; PNG output additionally requires `ResvgPngRenderer`.

## CLI

```bash
xlsx-chart-converter list report.xlsx --json
xlsx-chart-converter inspect report.xlsx chart-1 --data-mode chart-cache-first
xlsx-chart-converter export report.xlsx --chart all --format svg --out ./charts
xlsx-chart-converter export report.xlsx --chart all --format png --out ./charts --scale 2
```

The CLI is implemented for list, inspect, and export workflows. It uses the same core facade and renderer adapters as library callers.

## Current Coverage

Implemented:

- Safe ZIP/package access through the original OOXML package.
- Workbook, worksheet, chartsheet, drawing, anchor, and chart relationship discovery.
- Classic chart parsing for common 2D chart structures used by the current fixture corpus.
- A1 reference parsing and data resolution from chart caches, worksheet XML, and optional ExcelJS-like workbook providers.
- Theme color/style resolution, including F1 style cascade fields for fills, lines, text, markers, per-point styles, axes, legend, and data labels.
- ECharts SVG option mapping for supported chart IR.
- SVG sanitizing, ID prefixing, and accessibility metadata.
- PNG rasterization through resvg with scale, background, and font-file diagnostics.
- Apache POI corpus export script for visual QA.

Known limits:

- Legacy `.xls`, `.xlsb`, encrypted workbooks, VBA execution, external relationships, and full formula calculation are not supported.
- Unsupported OOXML chart/style features are surfaced as diagnostics and may degrade visually.
- Full pixel-perfect Excel/WPS parity is still a later fidelity phase; F1 covers style cascade and IR exposure, not final layout parity.

## Architecture

See [docs/architecture.md](docs/architecture.md) for package boundaries, runtime flow, milestone status, and the next fidelity/production hardening work.
