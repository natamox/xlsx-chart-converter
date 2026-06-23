# excel-chart Architecture

This document describes the current architecture of the repository after the core discovery, parsing, rendering, CLI, and F1 fidelity work. It is no longer an M0 scaffold: the public facade and adapters are functional, while deeper layout fidelity, broader chart coverage, and production hardening remain future phases.

## Goals

- Parse Excel OOXML charts from `.xlsx` and `.xlsm` workbooks in Node.js.
- Keep workbook/cell data access separate from chart OOXML parsing.
- Produce a stable Chart IR that can be inspected, cached, tested, and rendered by interchangeable adapters.
- Render supported charts to SVG and PNG without forcing parsing-only users to install renderer dependencies.
- Report unsupported chart features through structured diagnostics instead of silently dropping behavior.
- Maintain a corpus preview workflow so fidelity improvements can be checked against real workbook examples.

## Non-Goals

- No support for legacy `.xls`, `.xlsb`, encrypted workbooks, or VBA execution.
- No full Excel formula engine. Formula cells may use cached values or host-provided workbook data.
- No reliance on ExcelJS private APIs or workbook re-save behavior for chart extraction.
- No automatic access to external OOXML relationships.
- No claim of complete pixel parity with Excel/WPS yet; that belongs to later fidelity phases.

## Package Layout

```text
packages/
  core/      Public facade, package reader, workbook/drawing/chart parsing,
             data resolution, theme/style resolution, Chart IR, diagnostics.
  echarts/   SVG renderer adapter that maps Chart IR to ECharts server-side SVG.
  svg/       SVG sanitizing, ID prefixing, and accessibility post-processing.
  resvg/     PNG renderer adapter that rasterizes processed SVG with resvg.
  cli/       Command-line interface for list, inspect, and export workflows.
fixtures/   Apache POI workbook corpus and regression fixtures.
scripts/    Corpus export and local QA helpers.
```

Dependency direction:

```text
core      -> no renderer packages
echarts   -> core + svg + echarts
svg       -> standalone SVG post-processing helpers
resvg     -> core + @resvg/resvg-js
cli       -> core + echarts + svg + resvg
```

`@natamox/excel-chart-core` exposes renderer interfaces, but it does not depend on ECharts or resvg. Parsing and model-inspection callers can use `core` without renderer adapters.

## Runtime Flow

1. The facade opens a workbook source through `openWorkbook()` or `createChartEngine().open()`.
2. `PackageReader` indexes the original OOXML ZIP package with package limits.
3. Workbook, sheet, chartsheet, relationship, drawing, and anchor parsers locate chart parts by relationships rather than filename assumptions.
4. Classic chart parsers build a normalized Chart IR from chart XML, theme XML, style parts, workbook metadata, and resolved data series.
5. Data resolution chooses chart caches, worksheet XML/current workbook values, or host providers according to `dataMode`.
6. Renderers convert Chart IR to SVG. The ECharts adapter post-processes SVG by sanitizing output, prefixing IDs, and adding accessibility metadata.
7. PNG rendering rasterizes the SVG with explicit dimensions, scale, background, and font options.

## Public API

```ts
import { openWorkbook } from '@natamox/excel-chart-core';
import { EChartsSvgRenderer } from '@natamox/excel-chart-echarts';
import { ResvgPngRenderer } from '@natamox/excel-chart-resvg';

const workbook = await openWorkbook(
  { path: 'report.xlsx' },
  {
    renderer: new EChartsSvgRenderer(),
    pngRenderer: new ResvgPngRenderer()
  }
);

const charts = await workbook.listCharts();
const model = await workbook.getChartModel(charts[0].id, {
  dataMode: 'chart-cache-first'
});
const svg = await workbook.render(model.id, { format: 'svg' });
const png = await workbook.render(model.id, {
  format: 'png',
  scale: 2,
  background: '#fff'
});

await workbook.close();
```

Source forms:

- `Buffer`
- `Uint8Array`
- `{ path: string }`
- `{ path: string; workbook?: ExcelJsWorkbookLike }`
- `{ buffer: Buffer; workbook?: ExcelJsWorkbookLike }`

Data modes:

- `chart-cache-first`
- `cache-only`
- `exceljs-first`
- `exceljs-only`

## Core Boundaries

### Package and XML

- `package/*` owns ZIP package access, content types, relationships, target resolution, and package limits.
- `xml/*` owns namespace-aware tree parsing for the OOXML parts used by the current parser.
- External relationships are not followed; unsupported or missing relationships are reported through descriptors, errors, or diagnostics depending on the call site.

### Workbook and Drawing

- `workbook/*` maps workbook parts, sheets, chartsheets, shared strings, themes, and worksheet data providers.
- `drawing/*` resolves alternate content, anchors, dimensions, and chart object locations.
- Chart discovery follows workbook relationships to worksheets/chartsheets, then drawing relationships to chart parts.

### Chart Parsing

- `chart/classic/*` handles supported classic OOXML chart families and normalizes them into Chart IR.
- Unsupported chart families or features remain visible through `Diagnostic[]`.
- Parsing keeps chart descriptors separate from full model parsing so `listCharts()` stays cheaper than `getChartModel()`.

### Data and Theme

- `data/*` resolves A1 references, ranges, literals, chart caches, worksheet XML values, and optional ExcelJS-like providers.
- `theme/*` resolves theme colors, tint/shade transforms, fonts, chart style/color style parts, and F1 style cascade fields.
- Style IR currently includes shape fills/lines, text styles, markers, per-point styles, axes, legend, and data labels.

### IR and Diagnostics

- `public/types.ts` is the current public IR/type contract.
- `ChartModel.schemaVersion` is `1`.
- Diagnostics use stable codes, severities, messages, optional paths, and optional details.
- Unsupported fills such as gradient, pattern, or picture styles are exposed through diagnostics when degraded.

## Renderer Boundaries

### ECharts SVG

`packages/echarts` maps Chart IR to ECharts options and renders server-side SVG. It disables animation, applies explicit dimensions, maps number formats and labels, maps F1 style fields where possible, and post-processes SVG by default.

### SVG Post-Processing

`packages/svg` makes renderer output safer and easier to embed:

- remove scripts, foreign objects, event attributes, unsafe external links, and unsafe URL references;
- prefix IDs and `url(#id)` references to avoid collisions when multiple charts share a page;
- add `role="img"`, `aria-label`, and `<title>` metadata.

### resvg PNG

`packages/resvg` converts SVG to PNG. It centralizes scale, background, system font loading, custom font file loading, and font-file diagnostics.

## CLI Contract

```bash
excel-chart list report.xlsx --json
excel-chart inspect report.xlsx chart-1 --data-mode chart-cache-first
excel-chart export report.xlsx --chart all --format svg --out ./charts
excel-chart export report.xlsx --chart all --format png --out ./charts --scale 2
```

Exit behavior:

- usage and invalid flags return exit code `2`;
- no chart match returns exit code `4`;
- unexpected failures return exit code `6`.

## Corpus Preview

The repository includes Apache POI workbook fixtures under `fixtures/workbooks/apache-poi`.

```bash
make export-chart-corpus
```

This builds packages, exports every non-temporary `.xlsx` in the corpus, and writes:

- `output/chart-corpus-preview/index.html`
- `output/chart-corpus-preview/summary.json`
- per-chart `chart.svg`, `chart.png`, and `model.json`

`output/` is ignored by git.

## Milestone Status

| Area | Status |
| --- | --- |
| M0 Workspace scaffold | Implemented |
| M1 Package/workbook/drawing/chart discovery | Implemented |
| M2 Classic chart parsing and diagnostics | Implemented for current supported corpus |
| M3 Data resolution and worksheet providers | Implemented |
| M4 ECharts SVG rendering and SVG post-processing | Implemented |
| M5 resvg PNG rendering and CLI export | Implemented |
| F1 Style cascade and richer style IR | Implemented |
| F2 Layout fidelity | Not started |
| F3 High-fidelity SVG renderer | Not started |
| M6 Production hardening | Not started |

## Testing Strategy

Current checks:

```bash
pnpm run lint
pnpm run test
pnpm run typecheck
pnpm run build
pnpm run check
make export-chart-corpus
```

Coverage currently includes:

- facade lifecycle and closed-handle behavior;
- workbook/chart discovery through worksheet and chartsheet fixtures;
- A1 reference parsing and data resolution modes;
- anchor sizing and drawing coordinate conversion;
- theme/style/color resolution, including F1 cascade behavior;
- ECharts option mapping and renderer integration;
- SVG sanitizing, ID prefixing, and accessibility helpers;
- resvg PNG output and font diagnostics;
- Apache POI corpus discovery/parsing/export regression coverage.

Future fidelity work should add stable visual baselines and pixel comparison once layout behavior is intentionally frozen.

## Compatibility and Versioning

- Runtime target: Node.js >= 22.
- Module format: ESM with `NodeNext` TypeScript settings.
- Public APIs and IR schemas should be additive whenever possible.
- Breaking IR changes require a schema version bump and migration notes.
- Renderer dependency upgrades should run SVG and PNG regression tests plus the corpus preview because output can shift across ECharts or resvg versions.
