import type {
  ChartAxis,
  ChartAxisKind,
  ChartAxisPosition,
  ChartAxisStyle,
  ChartAxisScaling,
  ChartDataLabels,
  ChartDescriptor,
  ChartGroup,
  ChartGrouping,
  ChartLegend,
  ChartModel,
  ChartSeries,
  Diagnostic
} from '../../public/types.js';
import { attr, child, children, descendants, firstDescendant, parseXmlTree, textContent, valAttr } from '../../xml/xml-tree.js';

import { resolveChartData } from '../../data/data-resolver.js';
import { parseRelationships } from '../../package/relationship-parser.js';
import { relationshipPartName, resolveRelationshipTarget } from '../../package/relationship-resolver.js';
import type { PackageReader } from '../../package/package-reader.js';
import { resolveChartStyleTree } from '../../theme/style-resolver.js';
import { parseThemePart } from '../../theme/theme-parser.js';
import type { WorksheetDataProvider } from '../../workbook/worksheet-data-provider.js';
import type { WorkbookIndex } from '../../workbook/workbook-index.js';
import type { ChartDataProvider } from '../../data/data-resolver.js';
import type { XmlNode } from '../../xml/xml-tree.js';

export interface ParseClassicChartOptions {
  readonly packageReader: PackageReader;
  readonly descriptor: ChartDescriptor;
  readonly dataMode?: 'chart-cache-first' | 'exceljs-first' | 'cache-only' | 'exceljs-only';
  readonly fallbackDataProvider?: WorksheetDataProvider;
  readonly exceljsDataProvider?: ChartDataProvider;
  readonly workbookIndex?: WorkbookIndex;
}

const chartElementTypes = [
  'barChart',
  'lineChart',
  'areaChart',
  'pieChart',
  'doughnutChart',
  'scatterChart'
] as const;
const chartStyleRelType = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
const chartColorStyleRelType = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';

interface SeriesParseContext {
  readonly mode: 'chart-cache-first' | 'exceljs-first' | 'cache-only' | 'exceljs-only';
  readonly sheetName?: string;
  readonly fallbackDataProvider?: WorksheetDataProvider;
  readonly exceljsDataProvider?: ChartDataProvider;
}

interface ResolvedScalar {
  readonly value?: string | number | boolean;
  readonly diagnostics: readonly Diagnostic[];
}

interface ResolvedNumberValues {
  readonly values: readonly number[];
  readonly diagnostics: readonly Diagnostic[];
}

interface ResolvedCategoryValues {
  readonly values: readonly { label?: string; value?: number }[];
  readonly diagnostics: readonly Diagnostic[];
}

export async function parseClassicChartPart(
  options: ParseClassicChartOptions
): Promise<ChartModel> {
  if (!options.descriptor.chartPart) {
    throw new Error(`Chart descriptor ${options.descriptor.id} has no chart part.`);
  }

  const xml = (await options.packageReader.readPart(options.descriptor.chartPart)).toString('utf8');
  const theme = await readTheme(options.packageReader, options.workbookIndex?.themePart);
  const styleParts = await readChartStyleParts(options.packageReader, options.descriptor.chartPart);
  const chartSpace = parseXmlTree(xml, options.descriptor.chartPart);
  const chart = firstDescendant(chartSpace, 'chart') ?? chartSpace;
  const plotArea = firstDescendant(chart, 'plotArea') ?? chart;
  const chartGroups = chartElementTypes.flatMap((type) =>
    descendants(plotArea, type).map((node) => ({ type, node }))
  );
  const diagnostics: Diagnostic[] = [...options.descriptor.diagnostics];

  if (chartGroups.length === 0) {
    diagnostics.push({
      code: 'UNSUPPORTED_CHART_TYPE',
      severity: 'warning',
      message: 'No supported classic chart group was found.',
      path: options.descriptor.chartPart
    });
  }

  const series: ChartSeries[] = [];
  const dataDiagnostics: Diagnostic[] = [];
  for (const group of chartGroups) {
    const seriesContext = createSeriesParseContext(options);
    const groupSeries = await Promise.all(children(group.node, 'ser').map((ser, index) =>
      parseSeries(ser, group.type, parseAxisIds(group.node), parseDataLabels(group.node), index, seriesContext)
    ));
    for (const item of groupSeries) {
      series.push(item.series);
      dataDiagnostics.push(...item.diagnostics);
    }
  }

  const name = options.descriptor.name;
  const title = parseTitle(chart);
  const groups = chartGroups.map((group) => parseChartGroup(group.type, group.node));
  const styleResult = resolveChartStyleTree(chartSpace, theme, styleParts);
  const axes = parseAxes(plotArea, styleResult.style.axes);
  const legend = parseLegend(chart);
  const styledLegend = legend && styleResult.style.legend
    ? { ...legend, textStyle: styleResult.style.legend }
    : legend;
  return {
    schemaVersion: 1,
    id: options.descriptor.id,
    ...(name === undefined ? {} : { name }),
    ...(title === undefined ? {} : { title }),
    width: options.descriptor.width ?? 640,
    height: options.descriptor.height ?? 360,
    chartTypes: [...new Set(chartGroups.map((group) => normalizeChartType(group.type, group.node)))],
    ...(styledLegend ? { legend: styledLegend } : {}),
    plotArea: { chartGroups: groups },
    style: styleResult.style,
    axes,
    series,
    diagnostics: [...diagnostics, ...dataDiagnostics, ...styleResult.diagnostics]
  };
}

async function readTheme(packageReader: PackageReader, themePart: string | undefined) {
  if (!themePart || !(await packageReader.hasPart(themePart))) {
    return undefined;
  }
  return parseThemePart((await packageReader.readPart(themePart)).toString('utf8'), themePart);
}

async function readChartStyleParts(packageReader: PackageReader, chartPart: string) {
  const relsPart = relationshipPartName(chartPart);
  if (!(await packageReader.hasPart(relsPart))) {
    return {};
  }

  const rels = parseRelationships((await packageReader.readPart(relsPart)).toString('utf8'), relsPart);
  const chartStyle = rels.find((rel) =>
    rel.type === chartStyleRelType && rel.targetMode !== 'External'
  );
  const chartColorStyle = rels.find((rel) =>
    rel.type === chartColorStyleRelType && rel.targetMode !== 'External'
  );
  const styleXml = chartStyle
    ? await readOptionalRelatedPart(packageReader, chartPart, chartStyle.target)
    : undefined;
  const colorStyleXml = chartColorStyle
    ? await readOptionalRelatedPart(packageReader, chartPart, chartColorStyle.target)
    : undefined;
  return {
    ...(styleXml ? { chartStyleXml: styleXml } : {}),
    ...(colorStyleXml ? { chartColorStyleXml: colorStyleXml } : {})
  };
}

async function readOptionalRelatedPart(packageReader: PackageReader, sourcePart: string, target: string) {
  const partName = resolveRelationshipTarget(sourcePart, target);
  return (await packageReader.hasPart(partName))
    ? (await packageReader.readPart(partName)).toString('utf8')
    : undefined;
}

async function parseSeries(
  node: XmlNode,
  chartType: string,
  axisIds: readonly string[],
  groupDataLabels: ChartDataLabels | undefined,
  index: number,
  context: SeriesParseContext
): Promise<{ series: ChartSeries; diagnostics: Diagnostic[] }> {
  const nameResult = await parseSeriesName(node, context);
  const diagnostics = [...nameResult.diagnostics];
  const name = nameResult.value === undefined ? `Series ${index + 1}` : String(nameResult.value);
  const normalizedChartType = normalizeChartType(chartType, node);
  const dataLabels = parseDataLabels(node) ?? groupDataLabels;
  if (chartType === 'scatterChart') {
    const xs = await parseReferenceValues(child(node, 'xVal') ?? node, 'numCache', context);
    const ys = await parseReferenceNumberValues(child(node, 'yVal') ?? node, 'numCache', context);
    diagnostics.push(...xs.diagnostics, ...ys.diagnostics);
    return { series: {
      name,
      chartType: normalizedChartType,
      axisIds,
      ...(dataLabels ? { dataLabels } : {}),
      points: ys.values.map((y, pointIndex) => ({
        x: xs.values[pointIndex]?.value ?? xs.values[pointIndex]?.label ?? pointIndex + 1,
        y,
        value: y
      }))
    }, diagnostics };
  }

  const categories = await parseCategoryValues(child(node, 'cat') ?? node, context);
  const values = await parseReferenceNumberValues(child(node, 'val') ?? node, 'numCache', context);
  diagnostics.push(...categories.diagnostics, ...values.diagnostics);

  return { series: {
    name,
    chartType: normalizedChartType,
    axisIds,
    ...(dataLabels ? { dataLabels } : {}),
    points: values.values.map((value, pointIndex) => ({
      category: categories.values[pointIndex]?.label ?? String(pointIndex + 1),
      value,
      y: value
    }))
  }, diagnostics };
}

async function parseSeriesName(
  node: XmlNode,
  context: SeriesParseContext
): Promise<ResolvedScalar> {
  const tx = child(node, 'tx');
  if (!tx) {
    return { diagnostics: [] };
  }
  const formula = textContent(firstDescendant(tx, 'f'));
  const cacheValues = parseValues(firstDescendant(tx, 'strCache'))
    .concat(parseValues(firstDescendant(tx, 'numCache')))
    .flatMap((point) => point.label ?? point.value ?? []);
  const resolved = await resolveData(context, formula, cacheValues);
  const value = resolved.values[0];
  return value === undefined
    ? { diagnostics: resolved.diagnostics }
    : { value, diagnostics: resolved.diagnostics };
}

function parseTitle(chart: XmlNode): string | undefined {
  const title = child(chart, 'title');
  if (!title) {
    return undefined;
  }

  const cached = textContent(firstDescendant(title, 'strCache'))
    ?? textContent(firstDescendant(title, 'numCache'));
  if (cached) {
    return cached;
  }

  const rich = firstDescendant(title, 'rich');
  return textContent(rich);
}

function parseNumberValues(cache: XmlNode | undefined): number[] {
  return parseValues(cache).flatMap((point) => (
    typeof point.value === 'number' && Number.isFinite(point.value) ? [point.value] : []
  ));
}

async function parseReferenceNumberValues(
  node: XmlNode,
  cacheName: string,
  context: SeriesParseContext
): Promise<ResolvedNumberValues> {
  const cacheValues = parseNumberValues(firstDescendant(node, cacheName));
  const formula = textContent(firstDescendant(node, 'f'));
  const resolved = await resolveData(context, formula, cacheValues);
  return {
    values: resolved.values.flatMap((value) => {
    const numeric = typeof value === 'number' ? value : Number(value);
    return Number.isFinite(numeric) ? [numeric] : [];
    }),
    diagnostics: resolved.diagnostics
  };
}

async function parseReferenceValues(
  node: XmlNode,
  cacheName: string,
  context: SeriesParseContext
): Promise<ResolvedCategoryValues> {
  let cacheValues = parseValues(firstDescendant(node, cacheName));

  const multiLevel = firstDescendant(node, 'multiLvlStrCache');
  if (cacheValues.length === 0 && multiLevel) {
    cacheValues = parseMultiLevelStringValues(multiLevel);
  }

  const literal = parseValues(firstDescendant(node, cacheName === 'numCache' ? 'numLit' : 'strLit'));
  if (cacheValues.length === 0 && literal.length > 0) {
    cacheValues = literal;
  }

  const formula = textContent(firstDescendant(node, 'f'));
  const resolved = await resolveData(context, formula, cacheValues.map((point) => point.label ?? point.value ?? ''));
  return {
    values: resolved.values.map((value) => {
      const numeric = typeof value === 'number' ? value : Number(value);
      return {
        label: String(value),
        ...(Number.isFinite(numeric) ? { value: numeric } : {})
      };
    }),
    diagnostics: resolved.diagnostics
  };
}

async function parseCategoryValues(
  node: XmlNode,
  context: SeriesParseContext
): Promise<ResolvedCategoryValues> {
  if (firstDescendant(node, 'multiLvlStrRef') || firstDescendant(node, 'multiLvlStrCache')) {
    return parseReferenceValues(node, 'multiLvlStrCache', context);
  }
  if (firstDescendant(node, 'strRef') || firstDescendant(node, 'strCache') || firstDescendant(node, 'strLit')) {
    return parseReferenceValues(node, 'strCache', context);
  }
  return parseReferenceValues(node, 'numCache', context);
}

function parseValues(cache: XmlNode | undefined): { label?: string; value?: number }[] {
  if (!cache) {
    return [];
  }

  return children(cache, 'pt')
    .map((point) => {
      const text = textContent(child(point, 'v'));
      const numeric = text === undefined ? Number.NaN : Number(text);
      const value = Number.isFinite(numeric) ? numeric : undefined;
      return {
        index: Number(attr(point, 'idx') ?? '0'),
        label: text,
        ...(value === undefined ? {} : { value })
      };
    })
    .sort((left, right) => left.index - right.index)
    .map(({ label, value }) => ({
      ...(label === undefined ? {} : { label }),
      ...(value === undefined ? {} : { value })
    }));
}

function createSeriesParseContext(options: ParseClassicChartOptions): SeriesParseContext {
  const context: {
    mode: SeriesParseContext['mode'];
    sheetName?: string;
    fallbackDataProvider?: WorksheetDataProvider;
    exceljsDataProvider?: ChartDataProvider;
  } = {
    mode: options.dataMode ?? 'chart-cache-first'
  };
  if (options.descriptor.sheetName) {
    context.sheetName = options.descriptor.sheetName;
  }
  if (options.fallbackDataProvider) {
    context.fallbackDataProvider = options.fallbackDataProvider;
  }
  if (options.exceljsDataProvider) {
    context.exceljsDataProvider = options.exceljsDataProvider;
  }
  return context;
}

function dataProviderContext(context: SeriesParseContext) {
  return context.sheetName ? { sheetName: context.sheetName } : {};
}

function resolveData(
  context: SeriesParseContext,
  formula: string | undefined,
  cacheValues: readonly (string | number | boolean)[]
) {
  return resolveChartData({
    mode: context.mode,
    cacheValues,
    context: dataProviderContext(context),
    ...(formula ? { formula } : {}),
    ...(context.exceljsDataProvider ? { exceljsProvider: context.exceljsDataProvider } : {}),
    ...(context.fallbackDataProvider ? { fallbackProvider: context.fallbackDataProvider } : {})
  });
}

function normalizeChartType(type: string, node: XmlNode): string {
  if (type === 'barChart') {
    return valAttr(firstDescendant(node, 'barDir')) === 'bar' ? 'bar' : 'column';
  }
  return type.replace(/Chart$/, '').toLowerCase();
}

function parseLegend(chart: XmlNode): ChartLegend | undefined {
  const legend = child(chart, 'legend');
  if (!legend) {
    return undefined;
  }

  return {
    position: normalizeLegendPosition(valAttr(child(legend, 'legendPos'))),
    overlay: parseBooleanVal(child(legend, 'overlay')) ?? false
  };
}

function parseChartGroup(type: string, node: XmlNode): ChartGroup {
  const dataLabels = parseDataLabels(node);
  const grouping = parseGrouping(node);
  const varyColors = parseBooleanVal(child(node, 'varyColors'));
  const group: {
    type: string;
    axisIds: string[];
    grouping?: ChartGrouping;
    varyColors?: boolean;
    dataLabels?: ChartDataLabels;
  } = {
    type: normalizeChartType(type, node),
    axisIds: parseAxisIds(node)
  };
  if (grouping) {
    group.grouping = grouping;
  }
  if (varyColors !== undefined) {
    group.varyColors = varyColors;
  }
  if (dataLabels) {
    group.dataLabels = dataLabels;
  }
  return group;
}

function parseAxisIds(node: XmlNode): string[] {
  return children(node, 'axId').flatMap((axisId) => {
    const value = valAttr(axisId);
    return value ? [value] : [];
  });
}

function parseAxes(plotArea: XmlNode, styles: readonly { readonly axisId?: string }[] | undefined): ChartAxis[] {
  const axisNodes = children(plotArea, 'catAx')
    .concat(children(plotArea, 'dateAx'))
    .concat(children(plotArea, 'valAx'))
    .concat(children(plotArea, 'serAx'));
  return axisNodes.flatMap((axisNode) => {
    const id = valAttr(child(axisNode, 'axId'));
    if (!id) {
      return [];
    }

    const position = normalizeAxisPosition(valAttr(child(axisNode, 'axPos')));
    const title = parseTitle(axisNode);
    const crossesAxisId = valAttr(child(axisNode, 'crossAx'));
    const numberFormat = attr(child(axisNode, 'numFmt'), 'formatCode');
    const axis: {
      id: string;
      kind: ChartAxisKind;
      position?: ChartAxisPosition;
      title?: string;
      crossesAxisId?: string;
      numberFormat?: string;
      style?: ChartAxisStyle;
      scaling: ChartAxisScaling;
    } = {
      id,
      kind: axisKind(axisNode.localName),
      scaling: parseAxisScaling(axisNode)
    };
    if (position) {
      axis.position = position;
    }
    if (title) {
      axis.title = title;
    }
    if (crossesAxisId) {
      axis.crossesAxisId = crossesAxisId;
    }
    if (numberFormat) {
      axis.numberFormat = numberFormat;
    }
    const style = styles?.find((item) => item.axisId === id);
    if (style) {
      axis.style = style;
    }
    return [axis];
  });
}

function parseAxisScaling(axisNode: XmlNode): ChartAxisScaling {
  const scaling = child(axisNode, 'scaling');
  if (!scaling) {
    return {};
  }

  const orientation = valAttr(child(scaling, 'orientation'));
  const result: {
    orientation?: 'minMax' | 'maxMin';
    min?: number;
    max?: number;
    majorUnit?: number;
    minorUnit?: number;
    logBase?: number;
  } = {};
  if (orientation === 'maxMin' || orientation === 'minMax') {
    result.orientation = orientation;
  }
  const min = parseValNumber(child(scaling, 'min'));
  const max = parseValNumber(child(scaling, 'max'));
  const majorUnit = parseValNumber(child(axisNode, 'majorUnit'));
  const minorUnit = parseValNumber(child(axisNode, 'minorUnit'));
  const logBase = parseValNumber(child(scaling, 'logBase'));
  if (min !== undefined) {
    result.min = min;
  }
  if (max !== undefined) {
    result.max = max;
  }
  if (majorUnit !== undefined) {
    result.majorUnit = majorUnit;
  }
  if (minorUnit !== undefined) {
    result.minorUnit = minorUnit;
  }
  if (logBase !== undefined) {
    result.logBase = logBase;
  }
  return result;
}

function parseGrouping(node: XmlNode) {
  const value = valAttr(child(node, 'grouping'));
  if (value === 'clustered' || value === 'stacked' || value === 'percentStacked' || value === 'standard') {
    return value;
  }
  return undefined;
}

function parseDataLabels(node: XmlNode): ChartDataLabels | undefined {
  const labels = child(node, 'dLbls');
  if (!labels) {
    return undefined;
  }

  const result: {
    showValue?: boolean;
    showCategoryName?: boolean;
    showSeriesName?: boolean;
    showLegendKey?: boolean;
    showPercent?: boolean;
    showBubbleSize?: boolean;
    position?: string;
  } = {};
  assignBooleanLabel(result, 'showValue', parseBooleanVal(child(labels, 'showVal')));
  assignBooleanLabel(result, 'showCategoryName', parseBooleanVal(child(labels, 'showCatName')));
  assignBooleanLabel(result, 'showSeriesName', parseBooleanVal(child(labels, 'showSerName')));
  assignBooleanLabel(result, 'showLegendKey', parseBooleanVal(child(labels, 'showLegendKey')));
  assignBooleanLabel(result, 'showPercent', parseBooleanVal(child(labels, 'showPercent')));
  assignBooleanLabel(result, 'showBubbleSize', parseBooleanVal(child(labels, 'showBubbleSize')));
  const position = valAttr(child(labels, 'dLblPos'));
  if (position) {
    result.position = position;
  }
  return result;
}

function parseMultiLevelStringValues(cache: XmlNode): { label?: string; value?: number }[] {
  const levels = children(cache, 'lvl');
  const lastLevel = levels.at(-1);
  return parseValues(lastLevel);
}

function axisKind(localName: string) {
  if (localName === 'catAx') {
    return 'category';
  }
  if (localName === 'dateAx') {
    return 'date';
  }
  if (localName === 'serAx') {
    return 'series';
  }
  return 'value';
}

function normalizeAxisPosition(value: string | undefined) {
  if (value === 'l') {
    return 'left';
  }
  if (value === 'r') {
    return 'right';
  }
  if (value === 't') {
    return 'top';
  }
  if (value === 'b') {
    return 'bottom';
  }
  return undefined;
}

function normalizeLegendPosition(value: string | undefined): ChartLegend['position'] {
  if (value === 'l') {
    return 'left';
  }
  if (value === 'r') {
    return 'right';
  }
  if (value === 't') {
    return 'top';
  }
  if (value === 'b') {
    return 'bottom';
  }
  if (value === 'tr') {
    return 'corner';
  }
  return 'unknown';
}

function parseBooleanVal(node: XmlNode | undefined): boolean | undefined {
  const value = valAttr(node);
  if (value === undefined) {
    return undefined;
  }
  return value !== '0' && value !== 'false';
}

function parseValNumber(node: XmlNode | undefined): number | undefined {
  const value = Number(valAttr(node));
  return Number.isFinite(value) ? value : undefined;
}

function assignBooleanLabel(
  target: {
    showValue?: boolean;
    showCategoryName?: boolean;
    showSeriesName?: boolean;
    showLegendKey?: boolean;
    showPercent?: boolean;
    showBubbleSize?: boolean;
  },
  key: keyof Omit<ChartDataLabels, 'position'>,
  value: boolean | undefined
): void {
  if (value !== undefined) {
    target[key] = value;
  }
}
