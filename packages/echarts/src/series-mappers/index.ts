import { dataLabelOption } from '../labels.js';

import type {
  ChartDataPoint,
  ChartDataLabels,
  ChartGrouping,
  ChartMarkerStyle,
  ChartPointStyle,
  ChartShapeStyle
} from '@natamox/excel-chart-core';

export interface SeriesMapInput {
  readonly name: string;
  readonly chartType: string;
  readonly points: readonly ChartDataPoint[];
  readonly grouping?: ChartGrouping;
  readonly labels?: ChartDataLabels;
  readonly valueNumberFormat?: string;
  readonly xAxisIndex?: number;
  readonly yAxisIndex?: number;
  readonly percentStackedTotals?: readonly number[];
  readonly style?: ChartShapeStyle;
  readonly marker?: ChartMarkerStyle;
  readonly pointStyles?: readonly ChartPointStyle[];
}

export function mapCartesianSeries(input: SeriesMapInput): Record<string, unknown> {
  const type = seriesType(input.chartType);
  return {
    name: input.name,
    type,
    ...(input.xAxisIndex === undefined ? {} : { xAxisIndex: input.xAxisIndex }),
    ...(input.yAxisIndex === undefined ? {} : { yAxisIndex: input.yAxisIndex }),
    ...(input.chartType === 'area' ? { areaStyle: {} } : {}),
    ...(stackName(input.grouping) ? { stack: stackName(input.grouping) } : {}),
    ...(input.grouping === 'percentStacked' ? { stackStrategy: 'all' } : {}),
    ...itemStyle(input.style),
    ...markerStyle(input.marker),
    ...(type === 'scatter' ? {} : { connectNulls: false }),
    animation: false,
    label: dataLabelOption(input.labels, input.valueNumberFormat),
    data: type === 'scatter'
      ? input.points.map((point, index) => mapPointValueObject([toNumber(point.x), toNumber(point.y ?? point.value)], index, input))
      : input.points.map((point, index) => mapPointValue(point, index, input))
  };
}

export function mapPieSeries(input: SeriesMapInput & { readonly doughnut: boolean }): Record<string, unknown> {
  return {
    name: input.name,
    type: 'pie',
    radius: input.doughnut ? ['42%', '70%'] : '70%',
    animation: false,
    ...itemStyle(input.style),
    ...markerStyle(input.marker),
    label: dataLabelOption(input.labels, input.valueNumberFormat),
    data: input.points.map((point, index) => ({
      name: point.category ?? String(index + 1),
      value: point.value ?? point.y ?? null,
      ...pointItemStyle(index, input)
    }))
  };
}

function itemStyle(style: ChartShapeStyle | undefined): Record<string, unknown> {
  if (!style?.fill && !style?.line) {
    return {};
  }
  return {
    itemStyle: {
      ...(style.fill?.color ? { color: rgba(style.fill.transformedColor ?? style.fill.color, style.fill.alpha) } : {}),
      ...(style.fill?.kind === 'none' ? { color: 'transparent' } : {}),
      ...(style.line?.color && !style.line.noFill ? { borderColor: rgba(style.line.transformedColor ?? style.line.color, style.line.alpha) } : {}),
      ...(style.line?.width === undefined ? {} : { borderWidth: style.line.width })
    },
    ...(style.line?.color || style.line?.width || style.line?.dash
      ? { lineStyle: {
          ...(style.line.color && !style.line.noFill ? { color: rgba(style.line.transformedColor ?? style.line.color, style.line.alpha) } : {}),
          ...(style.line.width === undefined ? {} : { width: style.line.width }),
          ...(style.line.dash ? { type: mapDash(style.line.dash) } : {})
        } }
      : {})
  };
}

function markerStyle(marker: ChartMarkerStyle | undefined): Record<string, unknown> {
  if (!marker) {
    return {};
  }
  const shapeStyle: ChartShapeStyle = {
    ...(marker.fill ? { fill: marker.fill } : {}),
    ...(marker.line ? { line: marker.line } : {})
  };
  return {
    ...(marker.symbol ? { symbol: mapMarkerSymbol(marker.symbol) } : {}),
    ...(marker.size === undefined ? {} : { symbolSize: marker.size }),
    ...itemStyle(shapeStyle)
  };
}

function rgba(color: string, alpha: number | undefined): string {
  return alpha === undefined ? color : `${color}${Math.round(alpha * 255).toString(16).padStart(2, '0').toUpperCase()}`;
}

export function seriesType(chartType: string): 'bar' | 'line' | 'scatter' | 'pie' {
  if (chartType === 'line' || chartType === 'area') {
    return 'line';
  }
  if (chartType === 'scatter') {
    return 'scatter';
  }
  if (chartType === 'pie' || chartType === 'doughnut') {
    return 'pie';
  }
  return 'bar';
}

function stackName(grouping: ChartGrouping | undefined): string | undefined {
  return grouping === 'stacked' || grouping === 'percentStacked' ? 'total' : undefined;
}

function toNumber(value: unknown): number | null {
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(number) ? number : null;
}

function mapPointValue(point: ChartDataPoint, index: number, input: SeriesMapInput): number | null | Record<string, unknown> {
  const value = point.value ?? point.y;
  if (value === undefined) {
    return null;
  }
  const mappedValue = input.grouping !== 'percentStacked'
    ? value
    : normalizedPercentValue(value, index, input);
  return mapPointValueObject(mappedValue, index, input) as number | Record<string, unknown> | null;
}

function mapPointValueObject(
  value: unknown,
  index: number,
  input: SeriesMapInput
): unknown {
  const style = pointItemStyle(index, input);
  return Object.keys(style).length === 0 ? value : { value, ...style };
}

function pointItemStyle(index: number, input: SeriesMapInput): Record<string, unknown> {
  const pointStyle = input.pointStyles?.find((item) => item.index === index);
  const style = itemStyle(pointStyle?.style);
  const marker = markerStyle(pointStyle?.marker);
  return {
    ...style,
    ...marker
  };
}

function normalizedPercentValue(value: number, index: number, input: SeriesMapInput): number | null {
  const total = input.percentStackedTotals?.[index] ?? 0;
  return total === 0 ? null : value / total;
}

function mapDash(value: string): string {
  if (value === 'dash' || value === 'lgDash' || value === 'sysDash') {
    return 'dashed';
  }
  if (value === 'dot' || value === 'sysDot') {
    return 'dotted';
  }
  return value;
}

function mapMarkerSymbol(value: string): string {
  if (value === 'circle') {
    return 'circle';
  }
  if (value === 'square') {
    return 'rect';
  }
  if (value === 'diamond') {
    return 'diamond';
  }
  if (value === 'triangle') {
    return 'triangle';
  }
  if (value === 'none') {
    return 'none';
  }
  return value;
}
