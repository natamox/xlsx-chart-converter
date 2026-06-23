import { describe, expect, it } from 'vitest';
import { buildEChartsOption } from '../src/option-builder.js';

import type { ChartModel } from '@natamox/excel-chart-core';

type OptionAxis = Record<string, unknown> & {
  axisLabel?: { formatter: (value: number) => string };
};
type OptionSeries = Record<string, unknown> & {
  label?: { formatter: (item: { name?: string; seriesName?: string; value: number; percent?: number }) => string };
};

describe('buildEChartsOption', () => {
  it('maps combination series and secondary axis by chart group axis ids', () => {
    const option = buildEChartsOption(createComboModel()).option;

    expect(option.xAxis).toMatchObject([
      { id: 'cat-primary', type: 'category', data: ['Q1', 'Q2', 'Q3'] }
    ]);
    expect(option.yAxis).toMatchObject([
      { id: 'val-primary', position: 'left' },
      { id: 'val-secondary', position: 'right' }
    ]);
    expect(option.series).toMatchObject([
      { name: 'Revenue', type: 'bar', xAxisIndex: 0, yAxisIndex: 0 },
      { name: 'Margin', type: 'line', xAxisIndex: 0, yAxisIndex: 1 }
    ]);
  });

  it('maps scatter charts to numeric pairs without category axes', () => {
    const option = buildEChartsOption({
      ...baseModel,
      chartTypes: ['scatter'],
      axes: [
        { id: 'x', kind: 'value', position: 'bottom' },
        { id: 'y', kind: 'value', position: 'left' }
      ],
      plotArea: { chartGroups: [{ type: 'scatter', axisIds: ['x', 'y'] }] },
      series: [{
        name: 'Observations',
        chartType: 'scatter',
        axisIds: ['x', 'y'],
        points: [{ x: 1, y: 3 }, { x: 2, y: 8 }]
      }]
    }).option;

    expect(option.xAxis).toMatchObject([{ type: 'value' }]);
    expect(option.yAxis).toMatchObject([{ type: 'value' }]);
    expect(option.series).toMatchObject([{ type: 'scatter', data: [[1, 3], [2, 8]] }]);
  });

  it('maps scatter axis ids by series order when axis positions are absent', () => {
    const option = buildEChartsOption({
      ...baseModel,
      chartTypes: ['scatter'],
      axes: [
        { id: 'scatter-x', kind: 'value' },
        { id: 'scatter-y', kind: 'value' }
      ],
      plotArea: { chartGroups: [{ type: 'scatter', axisIds: ['scatter-x', 'scatter-y'] }] },
      series: [{
        name: 'Observations',
        chartType: 'scatter',
        axisIds: ['scatter-x', 'scatter-y'],
        points: [{ x: 1, y: 3 }]
      }]
    }).option;

    expect(option.xAxis).toMatchObject([{ id: 'scatter-x', type: 'value' }]);
    expect(option.yAxis).toMatchObject([{ id: 'scatter-y', type: 'value' }]);
    expect(option.series).toMatchObject([{ xAxisIndex: 0, yAxisIndex: 0 }]);
  });

  it('matches same-type chart groups by series axis ids', () => {
    const option = buildEChartsOption({
      ...baseModel,
      chartTypes: ['line'],
      axes: [
        { id: 'cat', kind: 'category', position: 'bottom' },
        { id: 'left', kind: 'value', position: 'left' },
        { id: 'right', kind: 'value', position: 'right' }
      ],
      plotArea: {
        chartGroups: [
          { type: 'line', axisIds: ['cat', 'left'], dataLabels: { showSeriesName: true } },
          { type: 'line', axisIds: ['cat', 'right'], dataLabels: { showValue: true } }
        ]
      },
      series: [
        { name: 'Left', chartType: 'line', axisIds: ['cat', 'left'], points: [{ category: 'A', value: 1 }] },
        { name: 'Right', chartType: 'line', axisIds: ['cat', 'right'], points: [{ category: 'A', value: 2 }] }
      ]
    }).option;

    const leftSeries = optionSeriesAt(option.series, 0);
    const rightSeries = optionSeriesAt(option.series, 1);
    expect(leftSeries.label?.formatter({ seriesName: 'Left', value: 1 })).toBe('Left');
    expect(rightSeries.label?.formatter({ value: 2 })).toBe('2');
  });

  it('applies labels, number formats, and blank point policy', () => {
    const option = buildEChartsOption({
      ...baseModel,
      axes: [
        { id: 'cat', kind: 'category', position: 'bottom' },
        { id: 'val', kind: 'value', position: 'left', numberFormat: '0.0%' }
      ],
      plotArea: {
        chartGroups: [{
          type: 'line',
          axisIds: ['cat', 'val'],
          dataLabels: { showValue: true, position: 'top' }
        }]
      },
      series: [{
        name: 'Rate',
        chartType: 'line',
        axisIds: ['cat', 'val'],
        points: [
          { category: 'A', value: 0.125 },
          { category: 'B' },
          { category: 'C', value: 0.25 }
        ]
      }]
    }).option;

    expect(option.series).toMatchObject([{
      label: { show: true, position: 'top' },
      data: [0.125, null, 0.25],
      connectNulls: false
    }]);
    const yAxis = firstOptionAxis(option.yAxis);
    const series = firstOptionSeries(option.series);
    expect(yAxis.axisLabel?.formatter).toEqual(expect.any(Function));
    expect(series.label?.formatter).toEqual(expect.any(Function));
    expect(yAxis.axisLabel?.formatter(0.125)).toBe('12.5%');
    expect(series.label?.formatter({ value: 0.125 })).toBe('12.5%');
  });

  it('maps pie and doughnut data labels', () => {
    const option = buildEChartsOption({
      ...baseModel,
      chartTypes: ['doughnut'],
      plotArea: {
        chartGroups: [{
          type: 'doughnut',
          axisIds: [],
          dataLabels: { showCategoryName: true, showValue: true }
        }]
      },
      series: [{
        name: 'Share',
        chartType: 'doughnut',
        points: [{ category: 'A', value: 4 }, { category: 'B', value: 6 }]
      }]
    }).option;

    expect(option.series).toMatchObject([{
      type: 'pie',
      radius: ['42%', '70%'],
      label: { show: true },
      data: [{ name: 'A', value: 4 }, { name: 'B', value: 6 }]
    }]);
    expect(firstOptionSeries(option.series).label?.formatter).toEqual(expect.any(Function));
  });

  it('formats pie percentage labels from ECharts percent payload', () => {
    const option = buildEChartsOption({
      ...baseModel,
      chartTypes: ['pie'],
      plotArea: {
        chartGroups: [{ type: 'pie', axisIds: [], dataLabels: { showCategoryName: true, showPercent: true } }]
      },
      series: [{
        name: 'Share',
        chartType: 'pie',
        points: [{ category: 'A', value: 4 }]
      }]
    }).option;

    const series = firstOptionSeries(option.series);
    expect(series.label?.formatter({ name: 'A', value: 4, percent: 40 })).toBe('A 40%');
  });

  it('normalizes percent stacked series to percentages', () => {
    const option = buildEChartsOption({
      ...baseModel,
      chartTypes: ['column'],
      axes: [
        { id: 'cat', kind: 'category', position: 'bottom' },
        { id: 'val', kind: 'value', position: 'left' }
      ],
      plotArea: {
        chartGroups: [{ type: 'column', axisIds: ['cat', 'val'], grouping: 'percentStacked' }]
      },
      series: [
        { name: 'A', chartType: 'column', axisIds: ['cat', 'val'], points: [{ category: 'Q1', value: 2 }, { category: 'Q2', value: 3 }] },
        { name: 'B', chartType: 'column', axisIds: ['cat', 'val'], points: [{ category: 'Q1', value: 6 }, { category: 'Q2', value: 3 }] }
      ]
    }).option;

    expect(option.yAxis).toMatchObject([{ max: 1 }]);
    expect(firstOptionAxis(option.yAxis).axisLabel?.formatter).toEqual(expect.any(Function));
    expect(option.series).toMatchObject([
      { stack: 'total', data: [0.25, 0.5] },
      { stack: 'total', data: [0.75, 0.5] }
    ]);
  });

  it('maps chart and series styles into ECharts option', () => {
    const option = buildEChartsOption({
      ...baseModel,
      chartTypes: ['column'],
      style: {
        chartArea: { fill: { color: '#F2F2F2' } },
        series: [{ fill: { color: '#4472C4', alpha: 0.5 }, line: { color: '#111111', width: 1.5 } }]
      },
      axes: [
        { id: 'cat', kind: 'category', position: 'bottom' },
        { id: 'val', kind: 'value', position: 'left' }
      ],
      plotArea: {
        chartGroups: [{ type: 'column', axisIds: ['cat', 'val'] }]
      },
      series: [
        { name: 'A', chartType: 'column', axisIds: ['cat', 'val'], points: [{ category: 'Q1', value: 2 }] }
      ]
    }).option;

    expect(option.backgroundColor).toBe('#F2F2F2');
    expect(option.series).toMatchObject([{
      itemStyle: { color: '#4472C480', borderColor: '#111111', borderWidth: 1.5 }
    }]);
  });

  it('maps F1 text, marker, dash, noFill, and per-point styles into ECharts option', () => {
    const option = buildEChartsOption({
      ...baseModel,
      title: 'Styled',
      chartTypes: ['line'],
      legend: { position: 'bottom', overlay: false, textStyle: { color: '#333333', fontSize: 9 } },
      style: {
        title: { color: '#222222', fontFamily: 'Aptos Display', fontSize: 14, bold: true },
        series: [{ fill: { kind: 'none' }, line: { color: '#ED7D31', width: 2, dash: 'dash' } }],
        seriesMarkers: [{ symbol: 'diamond', size: 7, fill: { color: '#00AA00' } }],
        pointStyles: [[{ index: 1, style: { fill: { color: '#FF0000' } } }]]
      },
      axes: [
        {
          id: 'cat',
          kind: 'category',
          position: 'bottom',
          style: { axisId: 'cat', text: { color: '#444444', italic: true } }
        },
        { id: 'val', kind: 'value', position: 'left' }
      ],
      plotArea: {
        chartGroups: [{ type: 'line', axisIds: ['cat', 'val'] }]
      },
      series: [{
        name: 'A',
        chartType: 'line',
        axisIds: ['cat', 'val'],
        points: [{ category: 'Q1', value: 2 }, { category: 'Q2', value: 4 }]
      }]
    }).option;

    expect(option.title).toMatchObject({
      textStyle: { color: '#222222', fontFamily: 'Aptos Display', fontSize: 14, fontWeight: 'bold' }
    });
    expect(option.legend).toMatchObject({
      bottom: 4,
      textStyle: { color: '#333333', fontSize: 9 }
    });
    expect(option.xAxis).toMatchObject([{
      axisLabel: { color: '#444444', fontStyle: 'italic' },
      nameTextStyle: { color: '#444444', fontStyle: 'italic' }
    }]);
    expect(option.series).toMatchObject([{
      symbol: 'diamond',
      symbolSize: 7,
      itemStyle: { color: '#00AA00' },
      lineStyle: { color: '#ED7D31', width: 2, type: 'dashed' },
      data: [2, { value: 4, itemStyle: { color: '#FF0000' } }]
    }]);
  });
});

const baseModel: ChartModel = {
  schemaVersion: 1,
  id: 'chart-1',
  title: 'Chart',
  width: 640,
  height: 360,
  chartTypes: ['column'],
  axes: [],
  series: [],
  diagnostics: []
};

function createComboModel(): ChartModel {
  return {
    ...baseModel,
    chartTypes: ['column', 'line'],
    axes: [
      { id: 'cat-primary', kind: 'category', position: 'bottom' },
      { id: 'val-primary', kind: 'value', position: 'left' },
      { id: 'val-secondary', kind: 'value', position: 'right', numberFormat: '0%' }
    ],
    plotArea: {
      chartGroups: [
        { type: 'column', axisIds: ['cat-primary', 'val-primary'], grouping: 'clustered' },
        { type: 'line', axisIds: ['cat-primary', 'val-secondary'], grouping: 'standard' }
      ]
    },
    series: [
      {
        name: 'Revenue',
        chartType: 'column',
        axisIds: ['cat-primary', 'val-primary'],
        points: [{ category: 'Q1', value: 10 }, { category: 'Q2', value: 15 }, { category: 'Q3', value: 13 }]
      },
      {
        name: 'Margin',
        chartType: 'line',
        axisIds: ['cat-primary', 'val-secondary'],
        points: [{ category: 'Q1', value: 0.2 }, { category: 'Q2', value: 0.24 }, { category: 'Q3', value: 0.22 }]
      }
    ]
  };
}

function firstOptionAxis(value: unknown): OptionAxis {
  return firstOptionObject(value);
}

function firstOptionSeries(value: unknown): OptionSeries {
  return optionSeriesAt(value, 0);
}

function optionSeriesAt(value: unknown, index: number): OptionSeries {
  return optionObjectAt(value, index);
}

function firstOptionObject(value: unknown): OptionAxis & OptionSeries {
  return optionObjectAt(value, 0);
}

function optionObjectAt(value: unknown, index: number): OptionAxis & OptionSeries {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error('Expected non-empty option array.');
  }
  const item: unknown = value[index];
  if (!item || typeof item !== 'object') {
    throw new Error('Expected option item object.');
  }
  return item as OptionAxis & OptionSeries;
}
