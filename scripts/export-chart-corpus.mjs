/* global console, process */
import fs from 'node:fs/promises';
import path from 'node:path';

import { openWorkbook } from '../packages/core/dist/index.js';
import { EChartsSvgRenderer } from '../packages/echarts/dist/index.js';
import { ResvgPngRenderer } from '../packages/resvg/dist/index.js';

const root = process.cwd();
const corpusDir = path.join(root, 'fixtures', 'workbooks', 'apache-poi');
const outDir = path.join(root, 'output', 'chart-corpus-preview');
const renderer = new EChartsSvgRenderer();
const pngRenderer = new ResvgPngRenderer();
const records = [];

await fs.rm(outDir, { recursive: true, force: true });
await fs.mkdir(outDir, { recursive: true });

const files = (await fs.readdir(corpusDir))
  .filter((name) => name.endsWith('.xlsx') && !name.startsWith('~$'))
  .sort();

for (const fileName of files) {
  const workbookOut = path.join(outDir, safeName(fileName.replace(/\.xlsx$/i, '')));
  await fs.mkdir(workbookOut, { recursive: true });
  const workbook = await openWorkbook({ path: path.join(corpusDir, fileName) }, { renderer, pngRenderer });
  try {
    const charts = await workbook.listCharts();
    for (const chart of charts) {
      const chartOut = path.join(workbookOut, safeName(chart.id));
      await fs.mkdir(chartOut, { recursive: true });
      try {
        const model = await workbook.getChartModel(chart.id);
        const svg = await workbook.render(chart.id, { format: 'svg' });
        const png = await workbook.render(chart.id, { format: 'png', scale: 2 });
        const svgPath = path.join(chartOut, 'chart.svg');
        const pngPath = path.join(chartOut, 'chart.png');
        const modelPath = path.join(chartOut, 'model.json');
        await fs.writeFile(svgPath, String(svg.data));
        await fs.writeFile(pngPath, png.data);
        await fs.writeFile(modelPath, JSON.stringify(model, null, 2));
        records.push({
          fileName,
          chartId: chart.id,
          chartName: chart.name,
          types: chart.chartTypes,
          width: svg.width,
          height: svg.height,
          diagnostics: model.diagnostics,
          svg: relative(outDir, svgPath),
          png: relative(outDir, pngPath),
          model: relative(outDir, modelPath)
        });
      } catch (error) {
        records.push({
          fileName,
          chartId: chart.id,
          chartName: chart.name,
          types: chart.chartTypes,
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }
    if (charts.length === 0) {
      records.push({ fileName, chartId: '', types: [], skipped: 'no charts' });
    }
  } finally {
    await workbook.close();
  }
}

await fs.writeFile(path.join(outDir, 'summary.json'), JSON.stringify(records, null, 2));
await fs.writeFile(path.join(outDir, 'index.html'), renderIndex(records));

const exported = records.filter((item) => item.png).length;
const failures = records.filter((item) => item.error).length;
const skipped = records.filter((item) => item.skipped).length;
console.log(JSON.stringify({ workbooks: files.length, charts: exported, failures, skipped, output: outDir }, null, 2));

function safeName(value) {
  return value.replace(/[^\w.-]+/g, '_').replace(/^_+|_+$/g, '') || 'item';
}

function relative(from, target) {
  return path.relative(from, target).replaceAll(path.sep, '/');
}

function renderIndex(items) {
  const cards = items.map((item) => {
    if (item.skipped) {
      return `<article class="card skipped"><h2>${escapeHtml(item.fileName)}</h2><p>No charts found.</p></article>`;
    }
    if (item.error) {
      return `<article class="card error"><h2>${escapeHtml(item.fileName)} / ${escapeHtml(item.chartId)}</h2><p>${escapeHtml(item.error)}</p></article>`;
    }
    const diagnostics = item.diagnostics?.length
      ? `<ul>${item.diagnostics.map((diag) => `<li>${escapeHtml(diag.severity)} ${escapeHtml(diag.code)}: ${escapeHtml(diag.message)}</li>`).join('')}</ul>`
      : '<p class="ok">No diagnostics</p>';
    return `<article class="card">
      <header>
        <h2>${escapeHtml(item.fileName)} / ${escapeHtml(item.chartId)}</h2>
        <p>${escapeHtml(item.types.join(', ') || 'unknown')} · ${item.width ?? ''}x${item.height ?? ''}</p>
      </header>
      <img src="${escapeHtml(item.png)}" alt="${escapeHtml(item.fileName)} ${escapeHtml(item.chartId)}">
      <nav><a href="${escapeHtml(item.svg)}">SVG</a><a href="${escapeHtml(item.png)}">PNG</a><a href="${escapeHtml(item.model)}">IR JSON</a></nav>
      ${diagnostics}
    </article>`;
  }).join('\n');
  return `<!doctype html>
<html lang="zh-CN">
<meta charset="utf-8">
<title>Excel Chart Corpus Preview</title>
<style>
body{font-family:Arial,"Microsoft YaHei",sans-serif;margin:0;background:#f6f7f9;color:#1f2937}
header.page{padding:24px 28px;background:#fff;border-bottom:1px solid #d9dee7}
h1{font-size:22px;margin:0 0 6px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:16px;padding:20px}
.card{background:#fff;border:1px solid #d9dee7;border-radius:8px;padding:14px}
.card h2{font-size:14px;margin:0 0 4px;word-break:break-all}
.card p{font-size:12px;margin:0 0 10px;color:#5b6472}
.card img{display:block;width:100%;height:auto;min-height:180px;border:1px solid #edf0f4;background:#fff}
.card nav{display:flex;gap:12px;margin:10px 0;font-size:12px}
.card ul{font-size:12px;margin:8px 0 0;padding-left:18px;color:#8a5a00}
.ok{color:#287347}
.error{border-color:#f1b6b6}
.skipped{opacity:.72}
</style>
<header class="page"><h1>Excel Chart Corpus Preview</h1><p>${items.filter((item) => item.png).length} charts exported · ${items.filter((item) => item.error).length} failures</p></header>
<main class="grid">${cards}</main>
</html>`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}
