# excel-chart Architecture

This document describes the target architecture for the repository as it exists today. The project is still an M0 scaffold: package boundaries, public facade types, build, lint, test, and CI are present, while chart discovery, parsing, and rendering are planned across later milestones.

## Goals

- Parse Excel OOXML charts from `.xlsx` and `.xlsm` workbooks in Node.js.
- Keep workbook and cell access separate from chart OOXML parsing.
- Produce a stable Chart IR that can be inspected, cached, tested, and rendered by interchangeable adapters.
- Render supported charts to SVG and PNG without making parsing-only users install renderer dependencies.
- Report unsupported chart features through structured diagnostics instead of silently dropping behavior.

## Non-Goals

- No support for legacy `.xls`, `.xlsb`, encrypted workbooks, or VBA execution.
- No full Excel formula engine. Formula cells may use cached values or host-provided data.
- No reliance on ExcelJS private APIs or workbook re-save behavior for chart extraction.
- No promise of pixel-perfect parity with desktop Excel in the first supported releases.
- No automatic access to external OOXML relationships.

## Package Layout

```text
packages/
  core/      Public facade, package reader, workbook/drawing/chart parsing,
             data resolution, theme resolution, Chart IR, diagnostics, limits.
  echarts/   SVG renderer adapter that maps Chart IR to ECharts server-side SVG.
  svg/       SVG sanitizing, ID prefixing, and accessibility post-processing.
  resvg/     PNG renderer adapter that rasterizes processed SVG with resvg.
  cli/       Command-line interface for list, inspect, and export workflows.
fixtures/   Workbook, IR, SVG, and PNG fixtures for regression tests.
```

The dependency direction is intentionally narrow:

```text
core      -> no renderer packages
echarts   -> core
svg       -> core
resvg     -> core
cli       -> core + echarts + svg + resvg
```

`@natamox/excel-chart-core` exposes renderer interfaces, but it does not depend on ECharts or resvg. This keeps parsing usable in services that only need chart metadata or IR.

## Runtime Flow

1. The facade opens a workbook source through `openWorkbook()` or `createChartEngine().open()`.
2. `PackageReader` indexes the original OOXML ZIP package with size and entry limits.
3. Workbook, sheet, relationship, drawing, and anchor parsers locate chart parts by relationships rather than filename assumptions.
4. Classic chart parsers build a normalized Chart IR from chart XML, theme XML, workbook metadata, and resolved data series.
5. Data resolution chooses chart caches, ExcelJS/current workbook values, or host providers according to `dataMode`.
6. Renderers convert Chart IR to SVG, then SVG post-processing sanitizes output, prefixes IDs, and adds accessibility metadata.
7. PNG rendering rasterizes the final SVG with explicit dimensions, scale, background, and font options.

## Public API Shape

The current facade already exposes the planned lifecycle:

```ts
import { openWorkbook } from '@natamox/excel-chart-core';

const workbook = await openWorkbook(buffer);
const charts = await workbook.listCharts();
const model = await workbook.getChartModel(charts[0].id);
const result = await workbook.render(model.id, { format: 'svg' });
await workbook.close();
```

In M0, `listCharts()` returns an empty list and model/render calls return structured `ERR_NOT_IMPLEMENTED` errors. Later milestones should preserve this lifecycle while filling in behavior.

## Core Boundaries

### Package and XML

- `package/*` owns ZIP package access, content types, relationships, target resolution, and package limits.
- `xml/*` owns secure namespace-aware parsing and XML limits.
- The parser must reject or diagnose unsafe constructs such as external relationships, DTD/entity usage, excessive nesting, oversized entries, and suspicious compression ratios.

### Workbook and Drawing

- `workbook/*` maps workbook parts, sheets, chartsheets, and defined names.
- `drawing/*` resolves drawing anchors and chart object dimensions.
- Chart discovery should follow workbook relationships to worksheets/chartsheets, then drawing relationships to chart parts.

### Chart Parsing

- `chart/classic/*` handles supported classic OOXML chart families such as bar, column, line, area, pie, doughnut, and scatter.
- `chart/chartex/*` detects modern ChartEx parts and returns diagnostics until dedicated support is implemented.
- Parsers should keep unsupported features visible in diagnostics and choose `warn`, `degrade`, or `error` through policy.

### Data and Theme

- `data/*` resolves A1 references, chart caches, literals, defined names, and external data providers.
- `theme/*` resolves theme colors, fonts, and style inheritance into explicit IR values.
- Data resolution must support cache-first and workbook-first modes because callers may render either the original file state or modified in-memory values.

### IR and Diagnostics

- `ir/*` owns schema versioning, normalization, and validation.
- `diagnostics/*` owns stable codes, severities, typed errors, and collectors.
- IR should be deterministic for the same input and options so snapshots remain useful.

## Renderer Boundaries

### ECharts SVG

`packages/echarts` maps Chart IR to ECharts options and renders server-side SVG. The adapter should use fixed dimensions, disabled animation, deterministic IDs where possible, and no workbook-provided JavaScript formatter functions.

### SVG Post-Processing

`packages/svg` makes renderer output safer and easier to embed:

- remove unsafe script, event, external reference, and foreign object content;
- prefix IDs and URL references to avoid collisions when multiple charts share a page;
- add title, description, role, and ARIA metadata when available.

### resvg PNG

`packages/resvg` converts the final SVG to PNG. It should centralize scale, background, font family, font file, and output dimension handling so CLI and library callers get the same rasterization behavior.

## CLI Contract

The CLI is the integration layer for common workflows:

```bash
excel-chart list report.xlsx --json
excel-chart inspect report.xlsx chart-1 --ir --diagnostics
excel-chart export report.xlsm --chart all --format png --out ./charts --scale 2
```

The package currently prints help and not-implemented messages. When implementation lands, CLI output should remain scriptable: JSON for inspection, stable exit codes, and diagnostics that identify unsupported charts or unsafe input.

## Milestones

| Milestone | Scope |
| --- | --- |
| M0 | Workspace scaffold, strict TypeScript, package boundaries, facade types, placeholder adapters, CI. |
| M1 | Safe package reader, content type and relationship parsing, workbook/sheet/drawing/chart discovery. |
| M2 | Classic chart parsing to Chart IR for common 2D chart families plus structured diagnostics. |
| M3 | Data resolution, theme/style resolution, IR validation, fixture snapshots. |
| M4 | ECharts SVG renderer and SVG post-processing integration. |
| M5 | resvg PNG renderer, CLI export, font/background/scale options. |
| M6 | Visual regression suite, performance limits, worker-pool rendering, broader chart coverage. |

## Testing Strategy

Tests should be requirement-driven:

- Facade lifecycle: open, list, model, render, close, and closed-handle errors.
- Discovery: worksheets, chartsheets, anchors, missing relationships, external relationships, malformed packages.
- Parsing: each supported chart family, empty series, mixed axes, labels, legends, and unsupported feature diagnostics.
- Data: cache-only, workbook-only, cache-first, workbook-first, literals, defined names, quoted sheet names, empty ranges, and formula cached values.
- Rendering: deterministic SVG snapshots, sanitized SVG, ID prefixing, accessibility metadata, PNG dimensions, scale, background, and fonts.
- Safety: package limits, XML limits, large point counts, text length limits, invalid input, and abort handling.

The repository already has Vitest wired at the workspace level. Add fixture-backed tests alongside the package that owns the behavior, and reserve pixel comparisons for stable renderer milestones.

## Compatibility and Versioning

- Runtime target: Node.js >= 22.
- Module format: ESM with `NodeNext` TypeScript settings.
- Public APIs and IR schemas should be additive whenever possible.
- Breaking IR changes require a schema version bump and migration notes.
- Renderer dependency upgrades should run SVG and PNG regression tests because output can shift across ECharts or resvg versions.
