# excel-chart

Node.js library for parsing Excel OOXML charts and rendering them to SVG or PNG. The project is organized as a pnpm workspace: `core` owns OOXML chart discovery, parsing, IR, diagnostics, and the public facade; ECharts, SVG post-processing, resvg PNG output, and the CLI evolve as separate packages so parsing-only users do not install rendering adapters by default.

This repository is currently in the M0 scaffolding phase. TypeScript, linting, tests, builds, CI, package boundaries, and public API skeletons are in place. Chart discovery, parsing, and rendering will land in later milestones.

## Requirements

- Node.js >= 22
- pnpm 11.5.1

## Quick Start

```bash
pnpm install
pnpm run check
```

Common commands:

```bash
pnpm run lint
pnpm run test
pnpm run typecheck
pnpm run build
make check
```

## Workspace Layout

```text
packages/
  core/              # Facade, PackageReader, workbook/drawing/chart/data/theme/xml/IR/diagnostics/runtime
  echarts/           # ECharts SSR SVG adapter
  svg/               # SVG sanitizing, ID prefixing, accessibility
  resvg/             # resvg PNG adapter
  cli/               # excel-chart CLI
fixtures/            # workbook and IR/SVG/PNG fixtures
```

Dependency direction:

```text
core <- exceljs adapter
core -> no ECharts/resvg dependency
echarts -> core IR types
svg -> core render types
resvg -> core render types
cli -> core + echarts + svg + resvg
```

## Public API Draft

```ts
import { openWorkbook } from '@natamox/excel-chart-core';

const workbook = await openWorkbook(buffer);
const charts = await workbook.listCharts();
await workbook.close();
```

Planned rendering API:

```ts
await workbook.render('chart-1', {
  format: 'png',
  scale: 2,
  dataMode: 'chart-cache-first'
});
```

## CLI Draft

```bash
excel-chart list report.xlsx --json
excel-chart inspect report.xlsx chart-1 --ir --diagnostics
excel-chart export report.xlsm --chart all --format png --out ./charts --scale 2
```

The current CLI only prints help and explicit not-implemented messages. Full inspection and export support will arrive in later milestones.

## Initial Boundaries

- ExcelJS is used only as the workbook and cell data adapter; private ExcelJS APIs are not used.
- The original `.xlsx` or `.xlsm` buffer is the source of truth for OOXML chart parts.
- `core` is decoupled from renderers; ECharts and resvg are replaceable adapters.
- `.xls`, `.xlsb`, encrypted workbooks, VBA execution, and full formula calculation are not supported.
