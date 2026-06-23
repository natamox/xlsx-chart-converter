#!/usr/bin/env node

import fs from 'node:fs/promises';
import path from 'node:path';

import { EChartsSvgRenderer } from '@natamox/excel-chart-echarts';
import { openWorkbook } from '@natamox/excel-chart-core';
import { ResvgPngRenderer } from '@natamox/excel-chart-resvg';
import { sanitizeSvg } from '@natamox/excel-chart-svg';

import type { DataMode, RenderFormat } from '@natamox/excel-chart-core';

const help = `excel-chart

Usage:
  excel-chart list <file> [--json]
  excel-chart inspect <file> <chart-id> [--data-mode <mode>]
  excel-chart export <file> --chart <id|name|all> --format <svg|png> --out <dir> [--scale <n>] [--data-mode <mode>]
`;

void main();

async function main(): Promise<void> {
  const [, , command, ...args] = process.argv;

  try {
    if (!command || command === '--help' || command === '-h') {
      process.stdout.write(help);
      return;
    }

    if (command === 'list') {
      await listCommand(args);
      return;
    }

    if (command === 'inspect') {
      await inspectCommand(args);
      return;
    }

    if (command === 'export') {
      await exportCommand(args);
      return;
    }

    throw new CliError(2, `Unknown command: ${command}.\n\n${help}`);
  } catch (error) {
    const cliError = error instanceof CliError ? error : new CliError(6, errorMessage(error));
    process.stderr.write(`${cliError.message}\n`);
    process.exitCode = cliError.exitCode;
  }
}

async function listCommand(args: string[]): Promise<void> {
  const file = args[0];
  if (!file) {
    throw new CliError(2, 'Usage: excel-chart list <file> [--json]');
  }

  const workbook = await openWorkbook({ path: file });
  try {
    const charts = await workbook.listCharts();
    if (args.includes('--json')) {
      process.stdout.write(`${JSON.stringify(charts, null, 2)}\n`);
      return;
    }
    for (const chart of charts) {
      process.stdout.write(`${chart.id}\t${chart.sheetName ?? ''}\t${chart.chartTypes.join(',')}\t${chart.name ?? ''}\n`);
    }
  } finally {
    await workbook.close();
  }
}

async function inspectCommand(args: string[]): Promise<void> {
  const [file, chartId] = args;
  if (!file || !chartId) {
    throw new CliError(2, 'Usage: excel-chart inspect <file> <chart-id> [--data-mode <mode>]');
  }
  const dataMode = parseDataMode(readFlag(args, '--data-mode'));

  const workbook = await openWorkbook({ path: file });
  try {
    const model = await workbook.getChartModel(chartId, dataMode ? { dataMode } : {});
    process.stdout.write(`${JSON.stringify(model, null, 2)}\n`);
  } finally {
    await workbook.close();
  }
}

async function exportCommand(args: string[]): Promise<void> {
  const file = args[0];
  if (!file) {
    throw new CliError(2, 'Usage: excel-chart export <file> --chart <id|name|all> --format <svg|png> --out <dir>');
  }

  const chartSelector = readFlag(args, '--chart') ?? 'all';
  const formatValue = readFlag(args, '--format') ?? 'svg';
  const outDir = readFlag(args, '--out');
  const scale = Number(readFlag(args, '--scale') ?? '1');
  const dataMode = parseDataMode(readFlag(args, '--data-mode'));

  if (!outDir || !isRenderFormat(formatValue)) {
    throw new CliError(2, 'Export requires --out and --format <svg|png>.');
  }
  const format = formatValue;

  await fs.mkdir(outDir, { recursive: true });
  const workbook = await openWorkbook({ path: file }, {
    renderer: new EChartsSvgRenderer(),
    pngRenderer: new ResvgPngRenderer()
  });

  try {
    const charts = await workbook.listCharts();
    const selected = chartSelector === 'all'
      ? charts
      : charts.filter((chart) => chart.id === chartSelector || chart.name === chartSelector);

    if (selected.length === 0) {
      throw new CliError(4, `No chart matched "${chartSelector}".`);
    }

    for (const chart of selected) {
      const result = await workbook.render(chart.id, {
        format,
        scale,
        ...(dataMode ? { dataMode } : {})
      });
      const data = format === 'svg' ? sanitizeSvg(String(result.data)) : result.data;
      const fileName = `${safeFileName(chart.name ?? chart.id)}.${format}`;
      const target = path.join(outDir, fileName);
      await fs.writeFile(target, data);
      process.stdout.write(`${target}\n`);
    }
  } finally {
    await workbook.close();
  }
}

function readFlag(args: string[], flag: string): string | undefined {
  const index = args.indexOf(flag);
  return index >= 0 ? args[index + 1] : undefined;
}

function isRenderFormat(value: string): value is RenderFormat {
  return value === 'svg' || value === 'png';
}

function parseDataMode(value: string | undefined): DataMode | undefined {
  if (!value) {
    return undefined;
  }
  if (value === 'chart-cache-first' || value === 'exceljs-first' || value === 'cache-only' || value === 'exceljs-only') {
    return value;
  }
  throw new CliError(2, `Unsupported data mode: ${value}.`);
}

function safeFileName(value: string): string {
  return value.replace(/[^\w.-]+/g, '_').replace(/^_+|_+$/g, '') || 'chart';
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

class CliError extends Error {
  constructor(
    readonly exitCode: number,
    message: string
  ) {
    super(message);
  }
}
