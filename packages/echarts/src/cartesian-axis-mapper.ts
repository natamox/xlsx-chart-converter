import { createNumberFormatter } from './number-format.js';

import type { ChartAxis, ChartSeries, ChartTextStyle } from '@natamox/xlsx-chart-converter-core';

export interface AxisMap {
  readonly xAxis: readonly Record<string, unknown>[];
  readonly yAxis: readonly Record<string, unknown>[];
  readonly axisIndexById: ReadonlyMap<string, AxisLocation>;
}

export interface AxisLocation {
  readonly dimension: 'x' | 'y';
  readonly index: number;
}

export function mapCartesianAxes(
  axes: readonly ChartAxis[],
  categories: readonly string[],
  categoryLabels: readonly (readonly string[] | undefined)[],
  scatter: boolean,
  axisIdPairs: readonly (readonly string[] | undefined)[] = [],
  percentAxisIds: ReadonlySet<string> = new Set(),
  series: readonly ChartSeries[] = []
): AxisMap {
  const inferred = inferAxisDimensions(axisIdPairs);
  const xCandidates = axes.filter((axis) => axisDimension(axis, scatter, inferred) === 'x');
  const yCandidates = axes.filter((axis) => axisDimension(axis, scatter, inferred) === 'y');
  const xAxes = xCandidates.length > 0 ? xCandidates : [defaultCategoryAxis()];
  const yAxes = yCandidates.length > 0 ? yCandidates : [defaultValueAxis()];
  const axisIndexById = new Map<string, AxisLocation>();

  xAxes.forEach((axis, index) => axisIndexById.set(axis.id, { dimension: 'x', index }));
  yAxes.forEach((axis, index) => axisIndexById.set(axis.id, { dimension: 'y', index }));

  return {
    xAxis: xAxes.map((axis) => axisOption(axis, scatter ? undefined : categories, categoryLabels, scatter, false, axisValueExtent(axis, 'x', series))),
    yAxis: yAxes.map((axis) => axisOption(axis, undefined, categoryLabels, true, percentAxisIds.has(axis.id), axisValueExtent(axis, 'y', series))),
    axisIndexById
  };
}

export function axisPairForSeries(
  axisIds: readonly string[] | undefined,
  axisIndexById: ReadonlyMap<string, AxisLocation>
): { readonly xAxisIndex: number; readonly yAxisIndex: number } {
  let xAxisIndex = 0;
  let yAxisIndex = 0;
  for (const axisId of axisIds ?? []) {
    const location = axisIndexById.get(axisId);
    if (location?.dimension === 'x') {
      xAxisIndex = location.index;
    }
    if (location?.dimension === 'y') {
      yAxisIndex = location.index;
    }
  }
  return { xAxisIndex, yAxisIndex };
}

function axisOption(
  axis: ChartAxis,
  data: readonly string[] | undefined,
  categoryLabels: readonly (readonly string[] | undefined)[],
  forceValue: boolean,
  forcePercent: boolean = false,
  inferredExtent?: AxisExtent
): Record<string, unknown> {
  const type = forceValue
    ? 'value'
    : axis.kind === 'date'
      ? 'time'
      : axis.kind === 'value'
        ? 'value'
        : 'category';
  const numberFormatter = createNumberFormatter(forcePercent ? '0%' : axis.numberFormat);
  const categoryFormatter = data ? categoryLabelFormatter(categoryLabels) : undefined;
  const formatter = numberFormatter ?? categoryFormatter;
  return {
    id: axis.id,
    type,
    ...(data ? { data: [...data] } : {}),
    ...(axis.position ? { position: axis.position } : {}),
    ...(axis.title ? { name: axis.title } : {}),
    ...(axis.scaling?.orientation === 'maxMin' ? { inverse: true } : {}),
    ...(axis.scaling?.min === undefined ? inferredExtent?.min === undefined ? {} : { min: inferredExtent.min } : { min: axis.scaling.min }),
    ...(axis.scaling?.max === undefined ? inferredExtent?.max === undefined ? {} : { max: inferredExtent.max } : { max: axis.scaling.max }),
    ...(axis.scaling?.majorUnit === undefined ? inferredExtent?.interval === undefined ? {} : { interval: inferredExtent.interval } : { interval: axis.scaling.majorUnit }),
    ...(forcePercent ? { min: 0, max: 1 } : {}),
    ...(formatter || axis.style?.text ? { axisLabel: {
      ...(formatter ? { formatter } : {}),
      ...textStyleFields(axis.style?.text)
    } } : {}),
    ...(axis.style?.text ? { nameTextStyle: textStyleFields(axis.style.text) } : {}),
    ...(axis.majorGridLine ? { splitLine: { show: true, lineStyle: lineStyleFields(axis.majorGridLine) } } : {}),
    ...(axis.minorGridLine ? { minorSplitLine: { show: true, lineStyle: lineStyleFields(axis.minorGridLine) } } : {}),
    ...(axis.style?.shape?.line ? { axisLine: { lineStyle: {
      ...lineStyleFields(axis.style.shape.line)
    } } } : {})
  };
}

interface AxisExtent {
  readonly min: number;
  readonly max: number;
  readonly interval: number;
}

function axisValueExtent(axis: ChartAxis, dimension: 'x' | 'y', series: readonly ChartSeries[]): AxisExtent | undefined {
  if (axis.kind !== 'value' || axis.scaling?.min !== undefined || axis.scaling?.max !== undefined) {
    return undefined;
  }
  const values = series.flatMap((item) => {
    if (dimension === 'x') {
      return item.points.flatMap((point) => numeric(point.x));
    }
    return item.points.flatMap((point) => numeric(point.y ?? point.value));
  });
  if (values.length === 0) {
    return undefined;
  }
  return niceExtent(Math.min(...values), Math.max(...values));
}

function numeric(value: unknown): number[] {
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(number) ? [number] : [];
}

function niceExtent(rawMin: number, rawMax: number): AxisExtent | undefined {
  if (!Number.isFinite(rawMin) || !Number.isFinite(rawMax)) {
    return undefined;
  }
  if (rawMin === rawMax) {
    const padding = rawMin === 0 ? 1 : Math.abs(rawMin) * 0.2;
    rawMin -= padding;
    rawMax += padding;
  }
  const includeZeroMin = rawMin >= 0 ? 0 : rawMin;
  const range = rawMax - includeZeroMin;
  const interval = niceStep(range / 5);
  const min = rawMin >= 0 ? 0 : Math.floor(rawMin / interval) * interval;
  const max = Math.ceil(rawMax / interval) * interval + interval;
  return { min: roundScale(min), max: roundScale(max), interval: roundScale(interval) };
}

function niceStep(value: number): number {
  const exponent = Math.floor(Math.log10(value));
  const magnitude = 10 ** exponent;
  const residual = value / magnitude;
  if (residual <= 1) {
    return magnitude;
  }
  if (residual <= 2) {
    return 2 * magnitude;
  }
  if (residual <= 5) {
    return 5 * magnitude;
  }
  return 10 * magnitude;
}

function roundScale(value: number): number {
  return Number(value.toPrecision(12));
}

function textStyleFields(style: ChartTextStyle | undefined): Record<string, unknown> {
  if (!style) {
    return {};
  }
  return {
    ...(style.fontFamily ? { fontFamily: style.fontFamily } : {}),
    ...(style.fontSize === undefined ? {} : { fontSize: style.fontSize }),
    ...(style.bold === undefined ? {} : { fontWeight: style.bold ? 'bold' : 'normal' }),
    ...(style.italic === undefined ? {} : { fontStyle: style.italic ? 'italic' : 'normal' }),
    ...(style.color ? { color: rgba(style.color, style.alpha) } : {})
  };
}

function lineStyleFields(style: NonNullable<ChartAxis['majorGridLine']>): Record<string, unknown> {
  return {
    ...(style.color && !style.noFill ? { color: rgba(style.transformedColor ?? style.color, style.alpha) } : {}),
    ...(style.width === undefined ? {} : { width: style.width }),
    ...(style.dash ? { type: mapDash(style.dash) } : {})
  };
}

function categoryLabelFormatter(categoryLabels: readonly (readonly string[] | undefined)[]): ((value: unknown, index?: number) => string) | undefined {
  if (categoryLabels.every((levels) => !levels || levels.length <= 1)) {
    return undefined;
  }
  return (value, index) => {
    const label = String(value);
    const levels = index === undefined ? undefined : categoryLabels[index];
    return levels && levels.length > 1 ? levels.join('\n') : label;
  };
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

function rgba(color: string, alpha: number | undefined): string {
  return alpha === undefined ? color : `${color}${Math.round(alpha * 255).toString(16).padStart(2, '0').toUpperCase()}`;
}

function axisDimension(
  axis: ChartAxis,
  scatter: boolean,
  inferred: ReadonlyMap<string, 'x' | 'y'>
): 'x' | 'y' {
  const inferredDimension = inferred.get(axis.id);
  if (inferredDimension) {
    return inferredDimension;
  }
  if (axis.position === 'bottom' || axis.position === 'top') {
    return 'x';
  }
  if (axis.position === 'left' || axis.position === 'right') {
    return 'y';
  }
  if (!scatter && (axis.kind === 'category' || axis.kind === 'date')) {
    return 'x';
  }
  return 'y';
}

function inferAxisDimensions(axisIdPairs: readonly (readonly string[] | undefined)[]): ReadonlyMap<string, 'x' | 'y'> {
  const dimensions = new Map<string, 'x' | 'y'>();
  for (const axisIds of axisIdPairs) {
    const [xAxisId, yAxisId] = axisIds ?? [];
    if (xAxisId && !dimensions.has(xAxisId)) {
      dimensions.set(xAxisId, 'x');
    }
    if (yAxisId && !dimensions.has(yAxisId)) {
      dimensions.set(yAxisId, 'y');
    }
  }
  return dimensions;
}

function defaultCategoryAxis(): ChartAxis {
  return { id: 'category-axis', kind: 'category', position: 'bottom' };
}

function defaultValueAxis(): ChartAxis {
  return { id: 'value-axis', kind: 'value', position: 'left' };
}
