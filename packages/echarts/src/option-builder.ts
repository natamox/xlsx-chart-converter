import type { ChartAxis, ChartGroup, ChartModel, ChartSeries, ChartTextStyle } from '@natamox/excel-chart-core';
import type { ECBasicOption } from 'echarts/types/dist/shared.js';
import { axisPairForSeries, mapCartesianAxes } from './cartesian-axis-mapper.js';
import { mapCartesianSeries, mapPieSeries } from './series-mappers/index.js';

export interface EChartsOptionBuildResult {
  readonly option: ECBasicOption;
}

export function buildEChartsOption(model: ChartModel): EChartsOptionBuildResult {
  const chartType = model.chartTypes[0] ?? 'bar';
  const isPie = chartType === 'pie' || chartType === 'doughnut';
  const option: ECBasicOption = isPie ? buildPieOption(model, chartType === 'doughnut') : buildCartesianOption(model);
  return { option };
}

function buildCartesianOption(model: ChartModel): ECBasicOption {
  const scatter = model.chartTypes.includes('scatter');
  const categories = collectCategories(model.series);
  const percentAxisIds = collectPercentStackedAxisIds(model);
  const axes = mapCartesianAxes(model.axes, categories, scatter, [
    ...model.series.map((series) => series.axisIds),
    ...(model.plotArea?.chartGroups.map((group) => group.axisIds) ?? [])
  ], percentAxisIds);
  const seriesGroups = model.series.map((series) => chartGroupForSeries(model, series));
  const percentStackedTotals = collectPercentStackedTotals(model.series, seriesGroups);

  return {
    title: titleOption(model),
    legend: legendOption(model),
    grid: gridOption(model),
    ...(model.style?.chartArea?.fill?.color ? { backgroundColor: model.style.chartArea.fill.color } : {}),
    xAxis: axes.xAxis,
    yAxis: axes.yAxis,
    series: model.series.map((series, index) => {
      const chartType = series.chartType ?? model.chartTypes[0] ?? 'bar';
      const axisPair = axisPairForSeries(series.axisIds, axes.axisIndexById);
      const group = seriesGroups[index];
      const valueAxis = valueAxisForSeries(series, model.axes);
      const labels = series.dataLabels ?? group?.dataLabels;
      return mapCartesianSeries({
        name: series.name,
        chartType,
        points: series.points,
        xAxisIndex: axisPair.xAxisIndex,
        yAxisIndex: axisPair.yAxisIndex,
        ...(group?.grouping ? { grouping: group.grouping } : {}),
        ...(labels ? { labels } : {}),
        ...(valueAxis?.numberFormat ? { valueNumberFormat: valueAxis.numberFormat } : {}),
        ...(model.style?.series?.[index] ? { style: model.style.series[index] } : {}),
        ...(model.style?.seriesMarkers?.[index] ? { marker: model.style.seriesMarkers[index] } : {}),
        ...(model.style?.pointStyles?.[index] ? { pointStyles: model.style.pointStyles[index] } : {}),
        ...(percentStackedTotals ? { percentStackedTotals } : {})
      });
    })
  };
}

function buildPieOption(model: ChartModel, doughnut: boolean): ECBasicOption {
  const firstSeries = model.series[0];
  const groupLabels = chartGroupForType(model, doughnut ? 'doughnut' : 'pie')?.dataLabels;
  const labels = firstSeries?.dataLabels ?? groupLabels;
  return {
    title: titleOption(model),
    legend: legendOption(model, true),
    ...(model.style?.chartArea?.fill?.color ? { backgroundColor: model.style.chartArea.fill.color } : {}),
    series: [
      mapPieSeries({
        name: firstSeries?.name ?? model.title ?? model.id,
        chartType: doughnut ? 'doughnut' : 'pie',
        points: firstSeries?.points ?? [],
        doughnut,
        ...(model.style?.series?.[0] ? { style: model.style.series[0] } : {}),
        ...(model.style?.seriesMarkers?.[0] ? { marker: model.style.seriesMarkers[0] } : {}),
        ...(model.style?.pointStyles?.[0] ? { pointStyles: model.style.pointStyles[0] } : {}),
        ...(labels ? { labels } : {})
      })
    ]
  };
}

function titleOption(model: ChartModel): Record<string, unknown> | undefined {
  return model.title ? {
    text: model.title,
    left: 'center',
    top: 8,
    ...textStyleOption(model.style?.title)
  } : undefined;
}

function legendOption(model: ChartModel, forceShow = false): Record<string, unknown> {
  const position = model.legend?.position;
  return {
    show: forceShow || model.series.length > 1 || Boolean(model.legend),
    ...(position === 'left' ? { left: 12, top: 'middle', orient: 'vertical' } : {}),
    ...(position === 'right' || position === undefined || position === 'unknown' ? {
      right: 12,
      top: 'middle',
      orient: 'vertical'
    } : {}),
    ...(position === 'top' ? { top: 32, left: 'center', orient: 'horizontal' } : {}),
    ...(position === 'bottom' ? { bottom: 4, left: 'center', orient: 'horizontal' } : {}),
    ...textStyleOption(model.legend?.textStyle ?? model.style?.legend)
  };
}

function gridOption(model: ChartModel): Record<string, unknown> {
  const hasRightAxis = model.axes.some((axis) => axis.position === 'right');
  return {
    left: 48,
    right: hasRightAxis ? 56 : 24,
    top: model.title ? 64 : 28,
    bottom: model.legend?.position === 'bottom' ? 68 : 48,
    containLabel: true
  };
}

function textStyleOption(style: ChartTextStyle | undefined): Record<string, unknown> {
  if (!style) {
    return {};
  }
  return {
    textStyle: {
      ...(style.fontFamily ? { fontFamily: style.fontFamily } : {}),
      ...(style.fontSize === undefined ? {} : { fontSize: style.fontSize }),
      ...(style.bold === undefined ? {} : { fontWeight: style.bold ? 'bold' : 'normal' }),
      ...(style.italic === undefined ? {} : { fontStyle: style.italic ? 'italic' : 'normal' }),
      ...(style.color ? { color: rgba(style.color, style.alpha) } : {})
    }
  };
}

function rgba(color: string, alpha: number | undefined): string {
  return alpha === undefined ? color : `${color}${Math.round(alpha * 255).toString(16).padStart(2, '0').toUpperCase()}`;
}

function collectCategories(series: readonly ChartSeries[]): string[] {
  const firstWithCategories = series.find((item) => item.points.some((point) => point.category !== undefined));
  return firstWithCategories?.points.map((point, index) =>
    point.category ?? String(point.x ?? index + 1)
  ) ?? [];
}

function chartGroupForSeries(model: ChartModel, series: ChartSeries): ChartGroup | undefined {
  const chartType = series.chartType ?? model.chartTypes[0];
  if (!chartType) {
    return undefined;
  }
  const groups = model.plotArea?.chartGroups.filter((item) => item.type === chartType) ?? [];
  return groups.find((group) => axesMatch(group.axisIds, series.axisIds)) ?? groups[0];
}

function chartGroupForType(model: ChartModel, chartType: string) {
  return model.plotArea?.chartGroups.find((item) => item.type === chartType);
}

function valueAxisForSeries(series: ChartSeries, axes: readonly ChartAxis[]): ChartAxis | undefined {
  const axisIds = series.axisIds ?? [];
  return axisIds
    .map((axisId) => axes.find((axis) => axis.id === axisId))
    .find((axis): axis is ChartAxis => axis?.kind === 'value');
}

function axesMatch(groupAxisIds: readonly string[], seriesAxisIds: readonly string[] | undefined): boolean {
  if (!seriesAxisIds || seriesAxisIds.length === 0) {
    return false;
  }
  return groupAxisIds.length === seriesAxisIds.length &&
    groupAxisIds.every((axisId, index) => axisId === seriesAxisIds[index]);
}

function collectPercentStackedTotals(
  series: readonly ChartSeries[],
  groups: readonly (ChartGroup | undefined)[]
): readonly number[] | undefined {
  const hasPercentStacked = groups.some((group) => group?.grouping === 'percentStacked');
  if (!hasPercentStacked) {
    return undefined;
  }

  const maxPoints = Math.max(0, ...series.map((item) => item.points.length));
  return Array.from({ length: maxPoints }, (_, pointIndex) =>
    series.reduce((total, item, seriesIndex) => {
      if (groups[seriesIndex]?.grouping !== 'percentStacked') {
        return total;
      }
      return total + Math.abs(item.points[pointIndex]?.value ?? item.points[pointIndex]?.y ?? 0);
    }, 0)
  );
}

function collectPercentStackedAxisIds(model: ChartModel): ReadonlySet<string> {
  const axisIds = new Set<string>();
  for (const group of model.plotArea?.chartGroups ?? []) {
    if (group.grouping === 'percentStacked') {
      const valueAxisId = group.axisIds[1];
      if (valueAxisId) {
        axisIds.add(valueAxisId);
      }
    }
  }
  return axisIds;
}
