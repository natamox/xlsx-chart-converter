import { createNumberFormatter } from './number-format.js';

import type { ChartAxis, ChartTextStyle } from '@natamox/excel-chart-core';

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
  scatter: boolean,
  axisIdPairs: readonly (readonly string[] | undefined)[] = [],
  percentAxisIds: ReadonlySet<string> = new Set()
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
    xAxis: xAxes.map((axis) => axisOption(axis, scatter ? undefined : categories, scatter)),
    yAxis: yAxes.map((axis) => axisOption(axis, undefined, true, percentAxisIds.has(axis.id))),
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
  forceValue: boolean,
  forcePercent: boolean = false
): Record<string, unknown> {
  const type = forceValue
    ? 'value'
    : axis.kind === 'date'
      ? 'time'
      : axis.kind === 'value'
        ? 'value'
        : 'category';
  const formatter = createNumberFormatter(forcePercent ? '0%' : axis.numberFormat);
  return {
    id: axis.id,
    type,
    ...(data ? { data: [...data] } : {}),
    ...(axis.position ? { position: axis.position } : {}),
    ...(axis.title ? { name: axis.title } : {}),
    ...(axis.scaling?.orientation === 'maxMin' ? { inverse: true } : {}),
    ...(axis.scaling?.min === undefined ? {} : { min: axis.scaling.min }),
    ...(axis.scaling?.max === undefined ? {} : { max: axis.scaling.max }),
    ...(axis.scaling?.majorUnit === undefined ? {} : { interval: axis.scaling.majorUnit }),
    ...(forcePercent ? { min: 0, max: 1 } : {}),
    ...(formatter || axis.style?.text ? { axisLabel: {
      ...(formatter ? { formatter } : {}),
      ...textStyleFields(axis.style?.text)
    } } : {}),
    ...(axis.style?.text ? { nameTextStyle: textStyleFields(axis.style.text) } : {}),
    ...(axis.style?.shape?.line ? { axisLine: { lineStyle: {
      ...(axis.style.shape.line.color ? { color: rgba(axis.style.shape.line.transformedColor ?? axis.style.shape.line.color, axis.style.shape.line.alpha) } : {}),
      ...(axis.style.shape.line.width === undefined ? {} : { width: axis.style.shape.line.width }),
      ...(axis.style.shape.line.dash ? { type: axis.style.shape.line.dash } : {})
    } } } : {})
  };
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
