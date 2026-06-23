import type { ChartDescriptor, Diagnostic } from '../../public/types.js';
import { resolveAnchorSize } from '../../drawing/anchor-size-resolver.js';
import { selectAlternateContent } from '../../drawing/alternate-content.js';
import type { AnchorMarker, AnchorModel } from '../../drawing/anchor-model.js';
import type { PackageReader } from '../../package/package-reader.js';
import { parseRelationships } from '../../package/relationship-parser.js';
import {
  relationshipPartName,
  resolveRelationshipTarget
} from '../../package/relationship-resolver.js';
import { createDefaultSheetMetrics, parseSheetMetrics } from '../../workbook/sheet-metrics.js';
import type { WorkbookSheetIndex } from '../../workbook/workbook-index.js';
import { attr, child, childrenAny, descendants, parseXmlTree, textContent } from '../../xml/xml-tree.js';
import type { XmlNode } from '../../xml/xml-tree.js';

const drawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const chartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const chartExRelationshipType = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';

export interface LocateChartsOptions {
  readonly packageReader: PackageReader;
  readonly sheet: WorkbookSheetIndex;
}

export async function locateChartsInDrawing(
  options: LocateChartsOptions
): Promise<readonly ChartDescriptor[]> {
  const sheetRelationships = await readRelationships(
    options.packageReader,
    relationshipPartName(options.sheet.partName)
  );
  const drawingRelationships = sheetRelationships.filter((relationship) =>
    relationship.type === drawingRelationshipType
  );

  const descriptors: ChartDescriptor[] = [];
  for (const drawingRelationship of drawingRelationships) {
    if (drawingRelationship.targetMode === 'External') {
      descriptors.push(unsupportedDescriptor({
        id: `${options.sheet.name}-external-drawing-${drawingRelationship.id}`,
        sheetName: options.sheet.name,
        diagnostics: [{
          code: 'EXTERNAL_RELATIONSHIP_SKIPPED',
          severity: 'warning',
          message: 'External drawing relationships are not resolved.',
          path: relationshipPartName(options.sheet.partName),
          details: { relationshipId: drawingRelationship.id, target: drawingRelationship.target }
        }]
      }));
      continue;
    }

    const drawingPart = resolveRelationshipTarget(options.sheet.partName, drawingRelationship.target);
    descriptors.push(...await locateChartsInDrawingPart(
      options.packageReader,
      options.sheet,
      drawingPart
    ));
  }

  return descriptors;
}

async function locateChartsInDrawingPart(
  packageReader: PackageReader,
  sheet: WorkbookSheetIndex,
  drawingPart: string,
  sourceDiagnostics: readonly Diagnostic[] = []
): Promise<ChartDescriptor[]> {
  if (!(await packageReader.hasPart(drawingPart))) {
    return [unsupportedDescriptor({
      id: stableMissingPartId(sheet.name, drawingPart),
      sheetName: sheet.name,
      drawingPart,
      diagnostics: [{
        code: 'MISSING_DRAWING_PART',
        severity: 'warning',
        message: 'Drawing relationship target was not found in the workbook package.',
        path: drawingPart
      }]
    })];
  }

  const drawingXml = (await packageReader.readPart(drawingPart)).toString('utf8');
  const drawing = selectAlternateContent(parseXmlTree(drawingXml, drawingPart));
  const drawingRelationships = await readRelationships(packageReader, relationshipPartName(drawingPart));
  const anchors = childrenAny(drawing, ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']);
  const sheetMetrics = sheet.kind === 'worksheet'
    ? await parseSheetMetrics(packageReader, sheet.partName)
    : createDefaultSheetMetrics();

  const descriptors: ChartDescriptor[] = [];
  for (const [index, anchor] of anchors.entries()) {
    const frame = descendants(anchor, 'graphicFrame')[0];
    if (!frame) {
      continue;
    }

    const chartRef = descendants(frame, 'chart')[0] ?? descendants(frame, 'chartEx')[0];
    const relId = attr(chartRef, 'r:id');
    if (!relId) {
      descriptors.push(unsupportedDescriptor({
        id: `${stableMissingPartId(sheet.name, drawingPart)}-missing-rel-id-${index + 1}`,
        name: attr(descendants(frame, 'cNvPr')[0], 'name') ?? `${sheet.name} Chart ${index + 1}`,
        sheetName: sheet.name,
        drawingPart,
        width: resolveAnchorSize(parseAnchor(anchor), sheetMetrics).width,
        height: resolveAnchorSize(parseAnchor(anchor), sheetMetrics).height,
        diagnostics: [{
          code: 'MISSING_CHART_RELATIONSHIP_ID',
          severity: 'warning',
          message: 'Chart frame does not contain a relationship id.',
          path: drawingPart
        }]
      }));
      continue;
    }

    const relationship = drawingRelationships.find((item) => item.id === relId);
    if (!relationship) {
      const size = resolveAnchorSize(parseAnchor(anchor), sheetMetrics);
      descriptors.push(unsupportedDescriptor({
        id: `${stableMissingPartId(sheet.name, drawingPart)}-missing-chart-rel-${relId}`,
        name: attr(descendants(frame, 'cNvPr')[0], 'name') ?? `${sheet.name} Chart ${index + 1}`,
        sheetName: sheet.name,
        drawingPart,
        width: size.width,
        height: size.height,
        diagnostics: [{
          code: 'MISSING_CHART_RELATIONSHIP',
          severity: 'warning',
          message: 'Drawing references a chart relationship that does not exist.',
          path: relationshipPartName(drawingPart),
          details: { relationshipId: relId }
        }]
      }));
      continue;
    }

    if (relationship.targetMode === 'External') {
      const size = resolveAnchorSize(parseAnchor(anchor), sheetMetrics);
      descriptors.push(unsupportedDescriptor({
        id: `${stableMissingPartId(sheet.name, drawingPart)}-external-chart-${relId}`,
        name: attr(descendants(frame, 'cNvPr')[0], 'name') ?? `${sheet.name} Chart ${index + 1}`,
        sheetName: sheet.name,
        drawingPart,
        width: size.width,
        height: size.height,
        diagnostics: [{
          code: 'EXTERNAL_RELATIONSHIP_SKIPPED',
          severity: 'warning',
          message: 'External chart relationships are not resolved.',
          path: relationshipPartName(drawingPart),
          details: { relationshipId: relId, target: relationship.target }
        }]
      }));
      continue;
    }

    const chartPart = resolveRelationshipTarget(drawingPart, relationship.target);
    const size = resolveAnchorSize(parseAnchor(anchor), sheetMetrics);
    const chartName = attr(descendants(frame, 'cNvPr')[0], 'name') ?? `${sheet.name} Chart ${index + 1}`;
    const chartTypes = relationship.type === chartRelationshipType
      ? await detectChartTypes(packageReader, chartPart)
      : relationship.type === chartExRelationshipType
      ? ['chartEx']
      : ['unknown'];
    const supported = relationship.type === chartRelationshipType;
    const chartPartExists = await packageReader.hasPart(chartPart);
    const diagnostics: Diagnostic[] = [...sourceDiagnostics];
    if (!supported) {
      diagnostics.push({
        code: 'UNSUPPORTED_CHART_EX',
        severity: 'warning',
        message: 'ChartEx charts are detected but not rendered by this version.',
        path: chartPart
      });
    }
    if (!chartPartExists) {
      diagnostics.push({
        code: 'MISSING_CHART_PART',
        severity: 'warning',
        message: 'Chart relationship target was not found in the workbook package.',
        path: chartPart
      });
    }

    descriptors.push({
      id: stableChartId(chartPart),
      name: chartName,
      sheetName: sheet.name,
      chartPart,
      drawingPart,
      width: size.width,
      height: size.height,
      chartTypes,
      supported: supported && chartPartExists,
      diagnostics
    });
  }

  return descriptors;
}

function unsupportedDescriptor(options: {
  id: string;
  name?: string;
  sheetName?: string;
  drawingPart?: string;
  width?: number;
  height?: number;
  diagnostics: Diagnostic[];
}): ChartDescriptor {
  return {
    id: sanitizeDescriptorId(options.id),
    ...(options.name === undefined ? {} : { name: options.name }),
    ...(options.sheetName === undefined ? {} : { sheetName: options.sheetName }),
    ...(options.drawingPart === undefined ? {} : { drawingPart: options.drawingPart }),
    ...(options.width === undefined ? {} : { width: options.width }),
    ...(options.height === undefined ? {} : { height: options.height }),
    chartTypes: ['unknown'],
    supported: false,
    diagnostics: options.diagnostics
  };
}

function parseAnchor(anchor: XmlNode): AnchorModel {
  if (anchor.localName === 'twoCellAnchor') {
    const from = parseMarker(child(anchor, 'from'));
    const to = parseMarker(child(anchor, 'to'));
    return {
      kind: 'twoCell',
      ...(from ? { from } : {}),
      ...(to ? { to } : {})
    };
  }

  const extNode = child(anchor, 'ext') ?? descendants(anchor, 'ext')[0];
  const ext = extNode ? {
    cxEmu: parseNumberAttr(extNode, 'cx'),
    cyEmu: parseNumberAttr(extNode, 'cy')
  } : undefined;
  const from = parseMarker(child(anchor, 'from'));

  return {
    kind: anchor.localName === 'oneCellAnchor' ? 'oneCell' : 'absolute',
    ...(from ? { from } : {}),
    ...(ext ? { ext } : {})
  };
}

function parseMarker(node: XmlNode | undefined): AnchorMarker | undefined {
  if (!node) {
    return undefined;
  }

  return {
    col: parseNumberText(child(node, 'col')),
    row: parseNumberText(child(node, 'row')),
    colOffEmu: parseNumberText(child(node, 'colOff')),
    rowOffEmu: parseNumberText(child(node, 'rowOff'))
  };
}

function parseNumberText(node: XmlNode | undefined): number {
  const parsed = Number(textContent(node) ?? '0');
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseNumberAttr(node: XmlNode, name: string): number {
  const parsed = Number(attr(node, name) ?? '0');
  return Number.isFinite(parsed) ? parsed : 0;
}

function stableChartId(chartPart: string): string {
  return chartPart
    .replace(/^xl\/charts\//, 'chart-')
    .replace(/\.xml$/i, '')
    .replace(/[^\w-]+/g, '-');
}

function stableMissingPartId(sheetName: string, partName: string): string {
  return sanitizeDescriptorId(`${sheetName}-${partName}`);
}

function sanitizeDescriptorId(value: string): string {
  return value.replace(/[^\w-]+/g, '-').replace(/^-+|-+$/g, '') || 'chart-unknown';
}

async function readRelationships(packageReader: PackageReader, relsPart: string) {
  if (!(await packageReader.hasPart(relsPart))) {
    return [];
  }
  return parseRelationships((await packageReader.readPart(relsPart)).toString('utf8'), relsPart);
}

async function detectChartTypes(packageReader: PackageReader, partName: string): Promise<string[]> {
  if (!(await packageReader.hasPart(partName))) {
    return ['unknown'];
  }

  const xml = (await packageReader.readPart(partName)).toString('utf8');
  const matches = [...xml.matchAll(/<c:(barChart|lineChart|areaChart|pieChart|doughnutChart|scatterChart)\b/g)]
    .map((match) => normalizeChartType(match[1] ?? 'unknown', xml));
  return [...new Set(matches)];
}

function normalizeChartType(type: string, xml: string): string {
  if (type === 'barChart') {
    return /<c:barDir\s+val="bar"/.test(xml) ? 'bar' : 'column';
  }
  return type.replace(/Chart$/, '').toLowerCase();
}
