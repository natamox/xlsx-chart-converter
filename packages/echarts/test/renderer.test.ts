import { describe, expect, it } from 'vitest';
import { EChartsSvgRenderer } from '../src/index.js';

import type { ChartModel } from '@natamox/xlsx-chart-converter-core';

describe('EChartsSvgRenderer', () => {
  it('renders sanitized accessible SVG through the SSR adapter', async () => {
    const renderer = new EChartsSvgRenderer({ idPrefix: 'm4-' });

    const svg = await renderer.render(model, {
      width: 480,
      height: 320,
      scale: 1,
      background: '#fff'
    });

    expect(svg).toContain('<svg');
    expect(svg).toContain('role="img"');
    expect(svg).toContain('aria-label="M4 Combo"');
    expect(svg).toContain('<title>M4 Combo</title>');
    expect(svg).not.toContain('<script');
    expect(svg).not.toContain('<foreignObject');
  });

  it('lets explicit render background override model style background', async () => {
    const renderer = new EChartsSvgRenderer({ postProcess: false });
    const svg = await renderer.render({
      ...model,
      style: { chartArea: { fill: { color: '#111111' } } }
    }, {
      width: 480,
      height: 320,
      scale: 1,
      background: '#ffffff'
    });

    expect(svg).toContain('fill="#ffffff"');
    expect(svg).not.toContain('fill="#111111"');
  });

  it('uses render context dimensions when inferring automatic value axis ticks', async () => {
    const renderer = new EChartsSvgRenderer({ postProcess: false });
    const svg = await renderer.render(lineModel, {
      width: 794,
      height: 365,
      scale: 1
    });

    expect(svg).toContain('>18</text>');
    expect(svg).not.toContain('>20</text>');
    expect(svg).not.toContain('>25</text>');
  });
});

const model: ChartModel = {
  schemaVersion: 1,
  id: 'm4-combo',
  title: 'M4 Combo',
  width: 480,
  height: 320,
  chartTypes: ['column', 'line'],
  legend: { position: 'bottom', overlay: false },
  axes: [
    { id: 'cat', kind: 'category', position: 'bottom' },
    { id: 'value', kind: 'value', position: 'left' },
    { id: 'rate', kind: 'value', position: 'right', numberFormat: '0%' }
  ],
  plotArea: {
    chartGroups: [
      { type: 'column', axisIds: ['cat', 'value'], grouping: 'clustered' },
      { type: 'line', axisIds: ['cat', 'rate'], grouping: 'standard', dataLabels: { showValue: true } }
    ]
  },
  series: [
    {
      name: 'Revenue',
      chartType: 'column',
      axisIds: ['cat', 'value'],
      points: [{ category: 'Q1', value: 12 }, { category: 'Q2', value: 18 }]
    },
    {
      name: 'Rate',
      chartType: 'line',
      axisIds: ['cat', 'rate'],
      points: [{ category: 'Q1', value: 0.2 }, { category: 'Q2', value: 0.25 }]
    }
  ],
  diagnostics: []
};

const lineModel: ChartModel = {
  schemaVersion: 1,
  id: 'line-chart',
  width: 538,
  height: 305,
  chartTypes: ['line'],
  axes: [
    { id: 'cat', kind: 'category', position: 'bottom' },
    { id: 'value', kind: 'value', position: 'left' }
  ],
  plotArea: {
    chartGroups: [{ type: 'line', axisIds: ['cat', 'value'] }]
  },
  series: [
    {
      name: '1st Column',
      chartType: 'line',
      axisIds: ['cat', 'value'],
      points: [1, 2, 3, 4, 5, 6].map((value) => ({ category: String(value), value }))
    },
    {
      name: '2nd Column',
      chartType: 'line',
      axisIds: ['cat', 'value'],
      points: [10, 12, 14, 9, 15, 17].map((value, index) => ({
        category: String(index + 1),
        value
      }))
    }
  ],
  diagnostics: []
};
