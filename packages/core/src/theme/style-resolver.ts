import { attr, child, children, descendants, firstDescendant, parseXmlTree, valAttr } from '../xml/xml-tree.js';
import { resolveDrawingColor } from './color-resolver.js';

import type {
  ChartAxisStyle,
  ChartFillStyle,
  ChartLineStyle,
  ChartMarkerStyle,
  ChartPointStyle,
  ChartShapeStyle,
  ChartStyle,
  ChartTextStyle,
  Diagnostic
} from '../public/types.js';
import type { ThemeModel } from './theme-parser.js';
import type { XmlNode } from '../xml/xml-tree.js';

const chartElementTypes = [
  'barChart',
  'lineChart',
  'areaChart',
  'pieChart',
  'doughnutChart',
  'scatterChart'
] as const;

export interface ChartStyleParts {
  readonly chartStyleXml?: string;
  readonly chartColorStyleXml?: string;
}

export interface ChartStyleResolveResult {
  readonly style: ChartStyle;
  readonly diagnostics: readonly Diagnostic[];
}

export function resolveChartStyle(xml: string, theme: ThemeModel | undefined, parts: ChartStyleParts = {}): ChartStyle {
  const chartSpace = parseXmlTree(xml, 'chart-style.xml');
  return resolveChartStyleTree(chartSpace, theme, parts).style;
}

export function resolveChartStyleTree(
  chartSpace: XmlNode,
  theme: ThemeModel | undefined,
  parts: ChartStyleParts = {}
): ChartStyleResolveResult {
  const diagnostics: Diagnostic[] = [];
  const defaults = resolveStyleDefaults(parts, theme, diagnostics);
  const chart = firstDescendant(chartSpace, 'chart') ?? chartSpace;
  const plotArea = firstDescendant(chart, 'plotArea');

  const seriesNodes = chartElementTypes.flatMap((type) =>
    descendants(plotArea ?? chart, type).flatMap((group) => descendants(group, 'ser'))
  );
  const series = seriesNodes.map((ser, index) =>
    mergeShapeStyles(defaults.series[index % Math.max(1, defaults.series.length)], resolveShapeStyle(child(ser, 'spPr'), theme, diagnostics))
  );
  const markers = seriesNodes.map((ser) => resolveMarkerStyle(child(ser, 'marker'), theme, diagnostics));
  const pointStyles = seriesNodes.map((ser) => resolvePointStyles(ser, theme, diagnostics));
  const axes = resolveAxisStyles(plotArea ?? chart, theme, diagnostics);

  const chartArea = presentStyle(mergeShapeStyles(
    defaults.chartArea,
    resolveShapeStyle(child(chartSpace, 'spPr') ?? child(chart, 'spPr'), theme, diagnostics)
  ));
  const plotAreaStyle = presentStyle(mergeShapeStyles(
    defaults.plotArea,
    resolveShapeStyle(child(plotArea ?? chart, 'spPr'), theme, diagnostics)
  ));
  const title = mergeTextStyles(defaults.title, resolveTextStyle(child(chart, 'title'), theme));
  const legend = mergeTextStyles(defaults.legend, resolveTextStyle(child(chart, 'legend'), theme));
  const dataLabels = resolveTextStyle(firstDescendant(plotArea ?? chart, 'dLbls'), theme);
  const style = {
    ...(chartArea ? { chartArea } : {}),
    ...(plotAreaStyle ? { plotArea: plotAreaStyle } : {}),
    ...(series.some((item) => Boolean(item.fill ?? item.line)) ? { series } : {}),
    ...(markers.some(Boolean) ? { seriesMarkers: markers } : {}),
    ...(pointStyles.some((item) => item.length > 0) ? { pointStyles } : {}),
    ...(title ? { title } : {}),
    ...(legend ? { legend } : {}),
    ...(axes.length > 0 ? { axes } : {}),
    ...(dataLabels ? { dataLabels } : {}),
    ...(theme?.fonts ? { fonts: theme.fonts } : {})
  };
  return {
    style,
    diagnostics
  };
}

function presentStyle(style: ChartShapeStyle): ChartShapeStyle | undefined {
  return style.fill || style.line ? style : undefined;
}

function resolveShapeStyle(
  node: XmlNode | undefined,
  theme: ThemeModel | undefined,
  diagnostics: Diagnostic[]
): ChartShapeStyle {
  const fill = resolveFillStyle(node, theme, diagnostics);
  const line = resolveLineStyle(child(node ?? emptyNode, 'ln'), theme);
  return {
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {})
  };
}

function resolveLineStyle(node: XmlNode | undefined, theme: ThemeModel | undefined): ChartLineStyle | undefined {
  if (!node) {
    return undefined;
  }
  const color = resolveDrawingColor(child(node, 'solidFill'), theme);
  const width = Number(node.attributes.get('w'));
  const dash = valAttr(child(node, 'prstDash'));
  const noFill = Boolean(child(node, 'noFill'));
  const line: ChartLineStyle = {
    ...(color?.color ? { color: color.color, ...(color.transformedColor ? { transformedColor: color.transformedColor } : {}) } : {}),
    ...(color?.alpha === undefined ? {} : { alpha: color.alpha }),
    ...(Number.isFinite(width) ? { width: width / 12700 } : {}),
    ...(dash ? { dash } : {}),
    ...(noFill ? { noFill } : {})
  };
  return line.color || line.width !== undefined || line.alpha !== undefined || line.dash || line.noFill ? line : undefined;
}

function resolveFillStyle(
  node: XmlNode | undefined,
  theme: ThemeModel | undefined,
  diagnostics: Diagnostic[]
): ChartFillStyle | undefined {
  if (!node) {
    return undefined;
  }
  if (child(node, 'noFill')) {
    return { kind: 'none' };
  }
  const solid = child(node, 'solidFill');
  if (solid) {
    const color = resolveDrawingColor(solid, theme);
    return color ? { kind: 'solid', ...color } : undefined;
  }
  const unsupported = [
    ['gradFill', 'gradient'],
    ['pattFill', 'pattern'],
    ['blipFill', 'picture']
  ] as const;
  for (const [localName, kind] of unsupported) {
    if (child(node, localName)) {
      diagnostics.push({
        code: 'UNSUPPORTED_FILL_STYLE',
        severity: 'warning',
        message: `${kind} fill is not fully supported and will be degraded.`,
        details: { kind }
      });
      return { kind };
    }
  }
  return undefined;
}

function resolveTextStyle(node: XmlNode | undefined, theme: ThemeModel | undefined): ChartTextStyle | undefined {
  if (!node) {
    return undefined;
  }
  const runProps = firstDescendant(node, 'defRPr') ?? firstDescendant(node, 'rPr');
  const latin = firstDescendant(runProps ?? node, 'latin');
  const color = resolveDrawingColor(child(runProps ?? emptyNode, 'solidFill'), theme);
  const size = Number(attr(runProps, 'sz'));
  const fontFamily = attr(latin, 'typeface');
  const bold = parseBooleanAttr(attr(runProps, 'b'));
  const italic = parseBooleanAttr(attr(runProps, 'i'));
  const style: ChartTextStyle = {
    ...(fontFamily ? { fontFamily } : {}),
    ...(Number.isFinite(size) ? { fontSize: size / 100 } : {}),
    ...(bold === undefined ? {} : { bold }),
    ...(italic === undefined ? {} : { italic }),
    ...(color?.color ? { color: color.transformedColor ?? color.color } : {}),
    ...(color?.alpha === undefined ? {} : { alpha: color.alpha })
  };
  return Object.keys(style).length > 0 ? style : undefined;
}

function resolveMarkerStyle(
  marker: XmlNode | undefined,
  theme: ThemeModel | undefined,
  diagnostics: Diagnostic[]
): ChartMarkerStyle | undefined {
  if (!marker) {
    return undefined;
  }
  const symbol = valAttr(child(marker, 'symbol'));
  const size = Number(valAttr(child(marker, 'size')));
  const shape = resolveShapeStyle(child(marker, 'spPr'), theme, diagnostics);
  const style: ChartMarkerStyle = {
    ...(symbol ? { symbol } : {}),
    ...(Number.isFinite(size) ? { size } : {}),
    ...(shape.fill ? { fill: shape.fill } : {}),
    ...(shape.line ? { line: shape.line } : {})
  };
  return Object.keys(style).length > 0 ? style : undefined;
}

function resolvePointStyles(
  ser: XmlNode,
  theme: ThemeModel | undefined,
  diagnostics: Diagnostic[]
): ChartPointStyle[] {
  return children(ser, 'dPt').flatMap((point) => {
    const index = Number(valAttr(child(point, 'idx')));
    if (!Number.isFinite(index)) {
      return [];
    }
    const style = presentStyle(resolveShapeStyle(child(point, 'spPr'), theme, diagnostics));
    const marker = resolveMarkerStyle(child(point, 'marker'), theme, diagnostics);
    return [{
      index,
      ...(style ? { style } : {}),
      ...(marker ? { marker } : {})
    }];
  });
}

function resolveAxisStyles(
  root: XmlNode,
  theme: ThemeModel | undefined,
  diagnostics: Diagnostic[]
): ChartAxisStyle[] {
  const axisNames = ['catAx', 'dateAx', 'valAx', 'serAx'];
  return axisNames.flatMap((name) =>
    children(root, name).flatMap((axis) => {
      const axisId = valAttr(child(axis, 'axId'));
      const shape = presentStyle(resolveShapeStyle(child(axis, 'spPr'), theme, diagnostics));
      const text = resolveTextStyle(axis, theme);
      if (!shape && !text) {
        return [];
      }
      return [{
        ...(axisId ? { axisId } : {}),
        ...(shape ? { shape } : {}),
        ...(text ? { text } : {})
      }];
    })
  );
}

function resolveStyleDefaults(
  parts: ChartStyleParts,
  theme: ThemeModel | undefined,
  diagnostics: Diagnostic[]
): {
  readonly chartArea?: ChartShapeStyle;
  readonly plotArea?: ChartShapeStyle;
  readonly title?: ChartTextStyle;
  readonly legend?: ChartTextStyle;
  readonly series: readonly ChartShapeStyle[];
} {
  const colorDefaults = parts.chartColorStyleXml
    ? parseColorStyleDefaults(parts.chartColorStyleXml, theme)
    : [];
  const chartStyle = parts.chartStyleXml
    ? parseXmlTree(parts.chartStyleXml, 'chart-style-part.xml')
    : undefined;
  if (!chartStyle) {
    return { series: colorDefaults.map((fill) => ({ fill })) };
  }

  const styleEntries = children(chartStyle, 'styleEntry');
  const seriesStyle = styleEntries.flatMap((entry, index) =>
    mergeShapeStyles(
      colorDefaults[index % Math.max(1, colorDefaults.length)],
      resolveShapeStyle(firstDescendant(entry, 'spPr'), theme, diagnostics)
    )
  );
  const title = styleEntries[2] ? resolveTextStyle(styleEntries[2], theme) : undefined;
  const legend = styleEntries[3] ? resolveTextStyle(styleEntries[3], theme) : undefined;
  return {
    ...(styleEntries[0] ? { chartArea: resolveShapeStyle(firstDescendant(styleEntries[0], 'spPr'), theme, diagnostics) } : {}),
    ...(styleEntries[1] ? { plotArea: resolveShapeStyle(firstDescendant(styleEntries[1], 'spPr'), theme, diagnostics) } : {}),
    ...(title ? { title } : {}),
    ...(legend ? { legend } : {}),
    series: seriesStyle.length > 0 ? seriesStyle : colorDefaults.map((fill) => ({ fill }))
  };
}

function parseColorStyleDefaults(xml: string, theme: ThemeModel | undefined): ChartFillStyle[] {
  const colorStyle = parseXmlTree(xml, 'chart-color-style-part.xml');
  return descendants(colorStyle, 'srgbClr')
    .concat(descendants(colorStyle, 'schemeClr'))
    .flatMap((node) => {
      const wrapper: XmlNode = {
        name: 'a:solidFill',
        localName: 'solidFill',
        attributes: new Map(),
        children: [node],
        text: ''
      };
      const color = resolveDrawingColor(wrapper, theme);
      return color ? [{ kind: 'solid' as const, ...color }] : [];
    });
}

function mergeShapeStyles(
  base: ChartShapeStyle | ChartFillStyle | undefined,
  override: ChartShapeStyle | undefined
): ChartShapeStyle {
  const baseShape = isFillStyle(base) ? { fill: base } : base;
  return {
    ...(baseShape?.fill || override?.fill ? { fill: mergeFillStyles(baseShape?.fill, override?.fill) } : {}),
    ...(baseShape?.line || override?.line ? { line: { ...baseShape?.line, ...override?.line } } : {})
  };
}

function mergeFillStyles(
  base: ChartFillStyle | undefined,
  override: ChartFillStyle | undefined
): ChartFillStyle {
  if (!override) {
    return { ...base };
  }
  if (override.kind && override.kind !== 'solid') {
    return { ...override };
  }
  return { ...base, ...override };
}

function mergeTextStyles(
  base: ChartTextStyle | undefined,
  override: ChartTextStyle | undefined
): ChartTextStyle | undefined {
  const style = { ...base, ...override };
  return Object.keys(style).length > 0 ? style : undefined;
}

function isFillStyle(value: ChartShapeStyle | ChartFillStyle | undefined): value is ChartFillStyle {
  return Boolean(value && ('kind' in value || 'color' in value || 'alpha' in value));
}

function parseBooleanAttr(value: string | undefined): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }
  return value !== '0' && value !== 'false';
}

const emptyNode: XmlNode = {
  name: '#empty',
  localName: '#empty',
  attributes: new Map(),
  children: [],
  text: ''
};
