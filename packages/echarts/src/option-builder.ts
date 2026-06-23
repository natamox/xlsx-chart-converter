import type { ChartAxis, ChartGroup, ChartManualLayout, ChartModel, ChartSeries, ChartShapeStyle, ChartTextStyle } from '@natamox/excel-chart-core';
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
  const categoryLabels = collectCategoryLabels(model.series);
  const percentAxisIds = collectPercentStackedAxisIds(model);
  const axes = mapCartesianAxes(model.axes, categories, categoryLabels, scatter, [
    ...model.series.map((series) => series.axisIds),
    ...(model.plotArea?.chartGroups.map((group) => group.axisIds) ?? [])
  ], percentAxisIds, model.series);
  const seriesGroups = model.series.map((series) => chartGroupForSeries(model, series));
  const percentStackedTotals = collectPercentStackedTotals(model.series, seriesGroups);

  return {
    title: titleOption(model),
    legend: legendOption(model),
    grid: gridOption(model),
    ...chartAreaOption(model),
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
        ...(group?.scatterStyle ? { scatterStyle: group.scatterStyle } : {}),
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
  const layout = pieLayoutOption(model, doughnut);
  return {
    title: titleOption(model),
    legend: legendOption(model, true),
    ...chartAreaOption(model),
    series: [
      mapPieSeries({
        name: firstSeries?.name ?? model.title ?? model.id,
        chartType: doughnut ? 'doughnut' : 'pie',
        points: firstSeries?.points ?? [],
        doughnut,
        ...(model.style?.series?.[0] ? { style: model.style.series[0] } : {}),
        ...(model.style?.seriesMarkers?.[0] ? { marker: model.style.seriesMarkers[0] } : {}),
        ...(model.style?.pointStyles?.[0] ? { pointStyles: model.style.pointStyles[0] } : {}),
        ...(labels ? { labels } : {}),
        ...layout
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
  const manualBox = absoluteLayoutBox(model.legend?.layout, model, legendPosition(model), { includeLegend: false });
  return {
    show: forceShow || model.series.length > 1 || Boolean(model.legend),
    ...(position === 'left' ? { left: 12, top: 'middle', orient: 'vertical' } : {}),
    ...(position === 'corner' ? { right: 12, top: 32, orient: 'vertical' } : {}),
    ...(position === 'right' || position === undefined || position === 'unknown' ? {
      right: 12,
      top: 'middle',
      orient: 'vertical'
    } : {}),
    ...(position === 'top' ? { top: 32, left: 'center', orient: 'horizontal' } : {}),
    ...(position === 'bottom' ? { bottom: 4, left: 'center', orient: 'horizontal' } : {}),
    ...(manualBox ? layoutBoxOption(manualBox) : {}),
    ...shapeStyleOption(model.legend?.style, 'legend'),
    ...textStyleOption(model.legend?.textStyle ?? model.style?.legend)
  };
}

function gridOption(model: ChartModel): Record<string, unknown> {
  const hasRightAxis = model.axes.some((axis) => axis.position === 'right');
  const legendBox = reservingLegendLayoutBox(model);
  const plotArea = absoluteLayoutBox(model.plotArea?.layout, model);
  return {
    left: plotArea ? plotArea.x : legendBox?.position === 'left' ? legendBox.right + 16 : 48,
    right: plotArea ? model.width - plotArea.right : legendBox?.position === 'right' || legendBox?.position === 'corner' ? model.width - legendBox.x + 16 : hasRightAxis ? 56 : 24,
    top: plotArea ? plotArea.y : legendBox?.position === 'top' ? legendBox.bottom + 16 : model.title ? 64 : 28,
    bottom: plotArea ? model.height - plotArea.bottom : legendBox?.position === 'bottom' ? model.height - legendBox.y + 16 : 48,
    containLabel: true,
    ...plotAreaStyleOption(model)
  };
}

function pieLayoutOption(model: ChartModel, doughnut: boolean): Record<string, unknown> {
  const plotArea = absoluteLayoutBox(model.plotArea?.layout, model) ?? inferredPlotAreaBox(model);
  if (!plotArea) {
    return {};
  }
  const radius = Math.max(16, Math.min(plotArea.width, plotArea.height) / 2);
  const outerRadius = `${Math.round((radius / Math.min(model.width, model.height)) * 100)}%`;
  return {
    center: [`${Math.round(plotArea.centerX / model.width * 100)}%`, `${Math.round(plotArea.centerY / model.height * 100)}%`],
    radius: doughnut ? [`${Math.round(Number.parseInt(outerRadius, 10) * 0.6)}%`, outerRadius] : outerRadius
  };
}

function inferredPlotAreaBox(model: ChartModel): LayoutBox | undefined {
  const legendBox = reservingLegendLayoutBox(model);
  if (!legendBox) {
    return undefined;
  }
  const left = legendBox.position === 'left' ? legendBox.right + 16 : 48;
  const right = legendBox.position === 'right' || legendBox.position === 'corner' ? legendBox.x - 16 : model.width - 24;
  const top = legendBox.position === 'top' ? legendBox.bottom + 16 : model.title ? 48 : 24;
  const bottom = legendBox.position === 'bottom' ? legendBox.y - 16 : model.height - 24;
  if (right <= left || bottom <= top) {
    return undefined;
  }
  return layoutBox(left, top, right - left, bottom - top, legendBox.position);
}

type ChartLegendPosition = NonNullable<ChartModel['legend']>['position'];

interface LayoutBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly right: number;
  readonly bottom: number;
  readonly centerX: number;
  readonly centerY: number;
  readonly position?: ChartLegendPosition;
}

interface LayoutReferenceOptions {
  readonly includeLegend?: boolean;
}

function legendLayoutBox(model: ChartModel): LayoutBox | undefined {
  if (!model.legend) {
    return undefined;
  }
  const manual = absoluteLayoutBox(model.legend.layout, model, legendPosition(model), { includeLegend: false });
  return manual ?? inferredLegendBox(model);
}

function reservingLegendLayoutBox(model: ChartModel): LayoutBox | undefined {
  if (model.legend?.overlay) {
    return undefined;
  }
  return legendLayoutBox(model);
}

function inferredLegendBox(model: ChartModel): LayoutBox | undefined {
  if (!model.legend || model.legend.overlay) {
    return undefined;
  }
  const position = legendPosition(model);
  const fontSize = model.legend.textStyle?.fontSize ?? model.style?.legend?.fontSize ?? 12;
  const itemCount = Math.max(1, model.series.length);
  const itemGap = Math.max(4, fontSize * 0.6);
  const itemHeight = Math.ceil(fontSize * 1.35);
  const itemWidth = Math.ceil(Math.max(...model.series.map((series) => textWidth(series.name, fontSize))) + fontSize * 2.8);
  const horizontalWidth = Math.min(model.width * 0.86, model.series.reduce((total, series) =>
    total + textWidth(series.name, fontSize) + fontSize * 3.4, 0));
  const horizontalHeight = itemHeight + itemGap;
  const width = position === 'left' || position === 'right'
    ? Math.min(model.width * 0.42, Math.max(itemWidth, model.width * 0.18))
    : horizontalWidth;
  const height = position === 'left' || position === 'right'
    ? Math.min(model.height * 0.82, itemCount * itemHeight + (itemCount - 1) * itemGap)
    : horizontalHeight;
  if (position === 'left') {
    return layoutBox(12, (model.height - height) / 2, width, height, position);
  }
  if (position === 'top') {
    return layoutBox((model.width - width) / 2, model.title ? 30 : 12, width, height, position);
  }
  if (position === 'bottom') {
    return layoutBox((model.width - width) / 2, model.height - height - 8, width, height, position);
  }
  return layoutBox(model.width - width - 12, (model.height - height) / 2, width, height, position);
}

function absoluteLayoutBox(
  layout: ChartManualLayout | undefined,
  model: ChartModel,
  position?: ChartLegendPosition,
  referenceOptions: LayoutReferenceOptions = {}
): LayoutBox | undefined {
  if (!layout || layout.x === undefined || layout.y === undefined || layout.width === undefined || layout.height === undefined) {
    return undefined;
  }
  const reference = layoutReferenceBox(layout, model, referenceOptions);
  const width = toAbsoluteLayoutValue(layout.width, layout.widthMode, reference.width);
  const height = toAbsoluteLayoutValue(layout.height, layout.heightMode, reference.height);
  const x = reference.x + toAbsoluteLayoutValue(layout.x, layout.xMode, reference.width);
  const y = reference.y + toAbsoluteLayoutValue(layout.y, layout.yMode, reference.height);
  if (width <= 0 || height <= 0) {
    return undefined;
  }
  return layoutBox(x, y, width, height, position);
}

function layoutReferenceBox(
  layout: ChartManualLayout,
  model: ChartModel,
  options: LayoutReferenceOptions = {}
): LayoutBox {
  if (layout.target === 'inner') {
    return innerLayoutReferenceBox(model, options);
  }
  return layoutBox(0, 0, model.width, model.height);
}

function innerLayoutReferenceBox(model: ChartModel, options: LayoutReferenceOptions = {}): LayoutBox {
  const axes = axisTextInsets(model);
  const labels = dataLabelInsets(model);
  const legend = options.includeLegend === false ? emptyInsets() : reservingLegendInsets(model);
  const left = 16 + axes.left + labels.left + legend.left;
  const right = 16 + axes.right + labels.right + legend.right;
  const top = chartTitleInset(model) + axes.top + labels.top + legend.top;
  const bottom = 20 + axes.bottom + labels.bottom + legend.bottom;
  return layoutBox(left, top, Math.max(1, model.width - left - right), Math.max(1, model.height - top - bottom));
}

function chartTitleInset(model: ChartModel): number {
  if (!model.title) {
    return 16;
  }
  const fontSize = model.style?.title?.fontSize ?? 14;
  return Math.ceil(8 + fontSize * 1.8);
}

function axisTextInsets(model: ChartModel): BoxInsets {
  const insets = emptyInsets();
  for (const axis of model.axes) {
    const fontSize = axis.style?.text?.fontSize ?? model.style?.axes?.find((item) => item.axisId === axis.id)?.text?.fontSize ?? 11;
    const labelInset = Math.ceil(fontSize * 1.8);
    const titleInset = axis.title ? Math.ceil((axis.style?.text?.fontSize ?? fontSize) * 2.2) : 0;
    const total = labelInset + titleInset;
    if (axis.position === 'left') {
      insets.left = Math.max(insets.left, total);
    } else if (axis.position === 'right') {
      insets.right = Math.max(insets.right, total);
    } else if (axis.position === 'top') {
      insets.top = Math.max(insets.top, total);
    } else {
      insets.bottom = Math.max(insets.bottom, total);
    }
  }
  return insets;
}

function dataLabelInsets(model: ChartModel): BoxInsets {
  const insets = emptyInsets();
  for (const group of model.plotArea?.chartGroups ?? []) {
    addDataLabelInset(insets, group.dataLabels, model);
  }
  for (const series of model.series) {
    addDataLabelInset(insets, series.dataLabels, model);
  }
  return insets;
}

function addDataLabelInset(insets: BoxInsets, labels: ChartSeries['dataLabels'], model: ChartModel): void {
  if (!labels || !hasVisibleDataLabel(labels)) {
    return;
  }
  const fontSize = model.style?.dataLabels?.fontSize ?? 11;
  const inset = Math.ceil(fontSize * 1.8);
  if (labels.position === 'outEnd' || labels.position === 'top') {
    insets.top = Math.max(insets.top, inset);
  } else if (labels.position === 'b' || labels.position === 'bottom') {
    insets.bottom = Math.max(insets.bottom, inset);
  } else if (labels.position === 'l' || labels.position === 'left') {
    insets.left = Math.max(insets.left, inset);
  } else if (labels.position === 'r' || labels.position === 'right') {
    insets.right = Math.max(insets.right, inset);
  }
}

function hasVisibleDataLabel(labels: NonNullable<ChartSeries['dataLabels']>): boolean {
  return Boolean(
    labels.showValue ||
    labels.showCategoryName ||
    labels.showSeriesName ||
    labels.showPercent ||
    labels.showBubbleSize ||
    labels.showLegendKey
  );
}

function reservingLegendInsets(model: ChartModel): BoxInsets {
  const legend = reservingLegendLayoutBox(model);
  const insets = emptyInsets();
  if (!legend) {
    return insets;
  }
  if (legend.position === 'left') {
    insets.left = legend.right + 16;
  } else if (legend.position === 'right' || legend.position === 'corner') {
    insets.right = model.width - legend.x + 16;
  } else if (legend.position === 'top') {
    insets.top = legend.bottom + 16;
  } else if (legend.position === 'bottom') {
    insets.bottom = model.height - legend.y + 16;
  }
  return insets;
}

interface BoxInsets {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

function emptyInsets(): BoxInsets {
  return { left: 0, right: 0, top: 0, bottom: 0 };
}

function toAbsoluteLayoutValue(value: number, mode: ChartManualLayout['xMode'], size: number): number {
  void mode;
  return value * size;
}

function layoutBox(x: number, y: number, width: number, height: number, position?: ChartLegendPosition): LayoutBox {
  return {
    x,
    y,
    width,
    height,
    right: x + width,
    bottom: y + height,
    centerX: x + width / 2,
    centerY: y + height / 2,
    ...(position ? { position } : {})
  };
}

function legendPosition(model: ChartModel): ChartLegendPosition {
  return model.legend?.position === 'unknown' ? 'right' : model.legend?.position ?? 'right';
}

function layoutBoxOption(box: LayoutBox): Record<string, unknown> {
  return {
    left: Math.round(box.x),
    top: Math.round(box.y),
    width: Math.round(box.width),
    height: Math.round(box.height)
  };
}

function textWidth(value: string, fontSize: number): number {
  return Array.from(value).reduce((total, char) => total + charWidth(char, fontSize), 0);
}

function charWidth(char: string, fontSize: number): number {
  if (/\s/u.test(char)) {
    return fontSize * 0.33;
  }
  if (/[\u4E00-\u9FFF\u3040-\u30FF\uAC00-\uD7AF]/u.test(char)) {
    return fontSize;
  }
  if (/[A-Z0-9]/u.test(char)) {
    return fontSize * 0.62;
  }
  return fontSize * 0.54;
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

function plotAreaStyleOption(model: ChartModel): Record<string, unknown> {
  return shapeStyleOption(model.style?.plotArea, 'grid');
}

function chartAreaOption(model: ChartModel): Record<string, unknown> {
  const style = model.style?.chartArea;
  if (!style?.fill && !style?.line) {
    return {};
  }
  return {
    ...(style.fill?.color ? { backgroundColor: rgba(style.fill.transformedColor ?? style.fill.color, style.fill.alpha) } : {}),
    ...(style.line?.color && !style.line.noFill ? {
      graphic: [{
        type: 'rect',
        left: 0,
        top: 0,
        shape: { width: model.width, height: model.height },
        style: {
          fill: 'transparent',
          stroke: rgba(style.line.transformedColor ?? style.line.color, style.line.alpha),
          lineWidth: style.line.width ?? 1
        },
        silent: true,
        z: -10
      }]
    } : {})
  };
}

function shapeStyleOption(style: ChartShapeStyle | undefined, target: 'grid' | 'legend'): Record<string, unknown> {
  if (!style?.fill && !style?.line) {
    return {};
  }
  const backgroundKey = target === 'grid' ? 'backgroundColor' : 'backgroundColor';
  return {
    ...(style.fill?.color ? { [backgroundKey]: rgba(style.fill.transformedColor ?? style.fill.color, style.fill.alpha) } : {}),
    ...(style.line?.color && !style.line.noFill ? { borderColor: rgba(style.line.transformedColor ?? style.line.color, style.line.alpha) } : {}),
    ...(style.line?.width === undefined ? {} : { borderWidth: style.line.width })
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

function collectCategoryLabels(series: readonly ChartSeries[]): readonly (readonly string[] | undefined)[] {
  const firstWithCategories = series.find((item) => item.points.some((point) => point.category !== undefined));
  return firstWithCategories?.points.map((point) => point.categoryLevels) ?? [];
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
