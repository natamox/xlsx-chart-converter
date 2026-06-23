import { createNumberFormatter } from './number-format.js';

import type { ChartDataLabels } from '@natamox/xlsx-chart-converter-core';

export function dataLabelOption(
  labels: ChartDataLabels | undefined,
  numberFormat: string | undefined
): Record<string, unknown> | undefined {
  if (!labels || !hasVisibleLabel(labels)) {
    return undefined;
  }

  const valueFormatter = createNumberFormatter(numberFormat);
  return {
    show: true,
    ...(labels.position ? { position: mapLabelPosition(labels.position) } : {}),
    formatter: (item: { name?: string; percent?: unknown; seriesName?: string; value?: unknown }) => {
      const parts: string[] = [];
      if (labels.showSeriesName && item.seriesName) {
        parts.push(item.seriesName);
      }
      if (labels.showCategoryName && item.name) {
        parts.push(item.name);
      }
      if (labels.showValue) {
        parts.push(valueFormatter ? valueFormatter(item.value) : safeLabelValue(item.value));
      }
      if (labels.showPercent) {
        parts.push(formatPercent(item.percent));
      }
      return parts.filter(Boolean).join(' ');
    }
  };
}

function formatPercent(value: unknown): string {
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(number) ? `${number.toFixed(0)}%` : '';
}

function safeLabelValue(value: unknown): string {
  if (value === undefined || value === null) {
    return '';
  }
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (Array.isArray(value)) {
    return value.map((item) => safeLabelValue(item)).filter(Boolean).join(', ');
  }
  return '';
}

function hasVisibleLabel(labels: ChartDataLabels): boolean {
  return Boolean(
    labels.showValue ||
    labels.showCategoryName ||
    labels.showSeriesName ||
    labels.showPercent ||
    labels.showBubbleSize ||
    labels.showLegendKey
  );
}

function mapLabelPosition(position: string): string {
  if (position === 'bestFit') {
    return 'inside';
  }
  if (position === 'ctr') {
    return 'inside';
  }
  if (position === 'inEnd') {
    return 'insideTop';
  }
  if (position === 'outEnd') {
    return 'top';
  }
  return position;
}
