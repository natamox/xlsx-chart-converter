# @natamox/excel-chart-cli

Command-line interface for inspecting and exporting Excel OOXML charts.

The CLI package wires together the core package, ECharts SVG rendering, SVG post-processing, and resvg PNG output.

## Usage

```bash
excel-chart list report.xlsx --json
excel-chart inspect report.xlsx chart-1 --ir --diagnostics
excel-chart export report.xlsm --chart all --format png --out ./charts --scale 2
```

## Status

The package is currently an M0 scaffold. The `excel-chart` binary prints help and explicit not-implemented messages until chart discovery and rendering land in later milestones.

## Runtime

Node.js >= 22.
