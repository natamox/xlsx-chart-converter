# @natamox/excel-chart-cli

Command-line interface for listing, inspecting, and exporting Excel OOXML charts.

The CLI wires together the core package, ECharts SVG rendering, SVG post-processing, and resvg PNG output.

## Usage

```bash
excel-chart list report.xlsx
excel-chart list report.xlsx --json
excel-chart inspect report.xlsx chart-1 --data-mode chart-cache-first
excel-chart export report.xlsx --chart all --format svg --out ./charts
excel-chart export report.xlsx --chart all --format png --out ./charts --scale 2
```

## Commands

- `list <file> [--json]`: discover charts and print descriptors.
- `inspect <file> <chart-id> [--data-mode <mode>]`: print the parsed `ChartModel` JSON.
- `export <file> --chart <id|name|all> --format <svg|png> --out <dir> [--scale <n>] [--data-mode <mode>]`: render selected charts.

Supported data modes:

- `chart-cache-first`
- `cache-only`
- `exceljs-first`
- `exceljs-only`

## Exit Codes

- `2`: usage error or unsupported flag value.
- `4`: export selector matched no charts.
- `6`: unexpected failure.

## Runtime

Node.js >= 22.
