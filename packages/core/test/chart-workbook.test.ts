import path from 'node:path';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { UnsupportedOperationError, createChartEngine, openWorkbook } from '../src/index.js';

import type { ExcelJsWorkbookLike, SvgRenderer } from '../src/index.js';

const fixturePath = path.resolve(import.meta.dirname, '../../../tmp/poi-WithThreeCharts.xlsx');
const chartSheetFixturePath = path.resolve(import.meta.dirname, '../../../tmp/exceljs-chart-sheet.xlsx');

describe('workbook facade', () => {
  it('lists charts from a real OOXML workbook', async () => {
    const workbook = await openWorkbook({ path: fixturePath });

    const charts = await workbook.listCharts();

    expect(charts).toHaveLength(3);
    expect(charts.every((chart) => chart.supported)).toBe(true);
    expect(charts.map((chart) => chart.id)).toEqual([
      'chart-chart1',
      'chart-chart2',
      'chart-chart3'
    ]);
    expect(charts.map((chart) => chart.name)).toEqual(['Chart 1', 'Chart 2', 'Chart 1']);
    expect(charts.map((chart) => `${chart.width}x${chart.height}`)).toEqual([
      '538x305',
      '532x288',
      '633x288'
    ]);
    expect(charts.map((chart) => chart.chartTypes[0])).toEqual(['line', 'pie', 'area']);
    expect(charts.map((chart) => chart.chartPart).sort()).toEqual([
      'xl/charts/chart1.xml',
      'xl/charts/chart2.xml',
      'xl/charts/chart3.xml'
    ]);
    await workbook.close();
  });

  it('lists charts from chartsheets with stable ids and anchor dimensions', async () => {
    const workbook = await openWorkbook({ path: chartSheetFixturePath });

    const charts = await workbook.listCharts();

    expect(charts).toMatchObject([{
      id: 'chart-chart1',
      name: 'Chart 1',
      sheetName: 'Chart1',
      chartPart: 'xl/charts/chart1.xml',
      drawingPart: 'xl/drawings/drawing1.xml',
      width: 977,
      height: 638,
      chartTypes: ['column'],
      supported: true
    }]);
    await workbook.close();
  });

  it('parses cached chart data into Chart IR', async () => {
    const workbook = await openWorkbook({ path: fixturePath });
    const [chart] = await workbook.listCharts();
    if (!chart) {
      throw new Error('Expected at least one chart in fixture.');
    }

    const model = await workbook.getChartModel(chart.id);

    expect(model.schemaVersion).toBe(1);
    expect(model.chartTypes.length).toBeGreaterThan(0);
    expect(model.plotArea?.chartGroups.length).toBeGreaterThan(0);
    expect(model.plotArea?.chartGroups[0]?.type).toBe('line');
    expect(model.plotArea?.chartGroups[0]?.axisIds.every((axisId) => typeof axisId === 'string')).toBe(true);
    expect(model.axes.map((axis) => axis.kind)).toEqual(expect.arrayContaining(['category', 'value']));
    expect(model.axes.some((axis) => axis.position)).toBe(true);
    expect(typeof model.legend?.position).toBe('string');
    expect(typeof model.legend?.overlay).toBe('boolean');
    expect(model.series.length).toBeGreaterThan(0);
    expect(model.series[0]?.chartType).toBe('line');
    expect(model.series[0]?.axisIds?.every((axisId) => typeof axisId === 'string')).toBe(true);
    expect(model.series[0]?.points.length).toBeGreaterThan(0);
    expect(toStableIrSnapshot(model)).toMatchInlineSnapshot(`
      {
        "axes": [
          {
            "kind": "category",
            "position": "bottom",
          },
          {
            "kind": "value",
            "position": "left",
          },
        ],
        "chartTypes": [
          "line",
        ],
        "groups": [
          {
            "axisIds": 2,
            "grouping": "standard",
            "type": "line",
          },
        ],
        "legend": {
          "overlay": false,
          "position": "right",
        },
        "series": [
          {
            "chartType": "line",
            "points": 6,
          },
          {
            "chartType": "line",
            "points": 6,
          },
        ],
      }
    `);
    await workbook.close();
  });

  it('preserves manual legend and plot area layout in Chart IR', async () => {
    const workbook = await openWorkbook(await createManualLayoutWorkbook());

    const [chart] = await workbook.listCharts();
    if (!chart) {
      throw new Error('Expected a chart in manual layout fixture.');
    }
    const model = await workbook.getChartModel(chart.id);

    await workbook.close();
    expect(model.legend?.layout).toEqual({
      target: 'inner',
      xMode: 'factor',
      yMode: 'factor',
      widthMode: 'factor',
      heightMode: 'factor',
      x: 0.71,
      y: 0.25,
      width: 0.24,
      height: 0.5
    });
    expect(model.plotArea?.layout).toEqual({
      target: 'inner',
      xMode: 'factor',
      yMode: 'factor',
      widthMode: 'factor',
      heightMode: 'factor',
      x: 0.08,
      y: 0.12,
      width: 0.58,
      height: 0.76
    });
  });

  it('returns diagnostics for damaged drawing and chart relationships', async () => {
    const workbook = await openWorkbook(await createDamagedRelationshipWorkbook());

    const charts = await workbook.listCharts();

    await workbook.close();
    expect(charts).toHaveLength(3);
    expect(charts.every((chart) => !chart.supported)).toBe(true);
    expect(charts.flatMap((chart) => chart.diagnostics.map((diagnostic) => diagnostic.code))).toEqual(
      expect.arrayContaining([
        'EXTERNAL_RELATIONSHIP_SKIPPED',
        'MISSING_DRAWING_PART',
        'MISSING_CHART_PART'
      ])
    );
  });

  it('uses ExcelJS data for exceljs-first without mutating cache-only output', async () => {
    const workbookLike = createWorkbookLike({
      Sheet1: {
        cells: {
          A1: 'M1',
          A2: 'M2',
          A3: 'M3',
          A4: 'M4',
          A5: 'M5',
          A6: 'M6',
          B1: 101,
          B2: 102,
          B3: 103,
          B4: 104,
          B5: 105,
          B6: 106
        }
      }
    });
    const workbook = await openWorkbook({ path: fixturePath, workbook: workbookLike });
    const [chart] = await workbook.listCharts();
    if (!chart) {
      throw new Error('Expected at least one chart in fixture.');
    }

    const cacheOnly = await workbook.getChartModel(chart.id, { dataMode: 'cache-only' });
    const exceljsFirst = await workbook.getChartModel(chart.id, { dataMode: 'exceljs-first' });

    await workbook.close();
    expect(cacheOnly.series[0]?.points.map((point) => point.value)).toEqual([1, 2, 3, 4, 5, 6]);
    expect(exceljsFirst.series[0]?.points).toEqual([]);
    expect(exceljsFirst.series[1]?.points.map((point) => point.category)).toEqual(['101', '102', '103', '104', '105', '106']);
    expect(exceljsFirst.series[1]?.points.map((point) => point.value)).toEqual([101, 102, 103, 104, 105, 106]);
  });

  it('filters hidden ExcelJS rows and columns in exceljs-only mode', async () => {
    const workbook = await openWorkbook({
      path: fixturePath,
      workbook: createWorkbookLike({
        Sheet1: {
          hiddenRows: new Set([2]),
          hiddenColumns: new Set([2]),
          cells: {
            A1: 'A',
            A2: 'B',
            A3: 'C',
            A4: 'D',
            A5: 'E',
            A6: 'F',
            B1: 11,
            B2: 12,
            B3: 13,
            B4: 14,
            B5: 15,
            B6: 16
          }
        }
      })
    });
    const [chart] = await workbook.listCharts();
    if (!chart) {
      throw new Error('Expected at least one chart in fixture.');
    }

    const model = await workbook.getChartModel(chart.id, { dataMode: 'exceljs-only' });

    await workbook.close();
    expect(model.series[0]?.points).toEqual([]);
    expect(model.diagnostics.map((diagnostic) => diagnostic.code)).toEqual(
      expect.arrayContaining(['DATA_REFERENCE_UNRESOLVED'])
    );
  });

  it('emits diagnostics when ExcelJS formula results may be stale', async () => {
    const workbook = await openWorkbook({
      path: fixturePath,
      workbook: createWorkbookLike({
        Sheet1: {
          cells: {
            A1: 'A',
            A2: 'B',
            A3: 'C',
            A4: 'D',
            A5: 'E',
            A6: 'F',
            B1: { result: 201 },
            B2: { result: 202 },
            B3: { result: 203 },
            B4: { result: 204 },
            B5: { result: 205 },
            B6: { result: 206 }
          }
        }
      })
    });
    const [chart] = await workbook.listCharts();
    if (!chart) {
      throw new Error('Expected at least one chart in fixture.');
    }

    const model = await workbook.getChartModel(chart.id, { dataMode: 'exceljs-first' });

    await workbook.close();
    expect(model.series[1]?.points[0]?.value).toBe(201);
    expect(model.diagnostics.map((diagnostic) => diagnostic.code)).toContain('STALE_FORMULA_RESULT_POSSIBLE');
  });

  it('renders SVG through an injected renderer', async () => {
    const renderer: SvgRenderer = {
      render: (model, context) => Promise.resolve(
        `<svg width="${context.width}" height="${context.height}"><title>${model.id}</title></svg>`
      )
    };
    const workbook = await createChartEngine({ renderer }).open({ path: fixturePath });
    const [chart] = await workbook.listCharts();
    if (!chart) {
      throw new Error('Expected at least one chart in fixture.');
    }

    const result = await workbook.render(chart.id, { format: 'svg', width: 320, height: 180 });

    expect(result.mediaType).toBe('image/svg+xml');
    expect(result.data).toContain('<svg');
    expect(result.width).toBe(320);
    expect(result.height).toBe(180);
    await workbook.close();
  });

  it('rejects invalid or unsupported package inputs', async () => {
    await expect(openWorkbook(Buffer.from([]))).rejects.toMatchObject({
      code: 'ERR_INVALID_PACKAGE'
    });
  });

  it('rejects unknown chart ids with a structured error', async () => {
    const workbook = await openWorkbook({ path: fixturePath });

    await expect(workbook.getChartModel('missing')).rejects.toMatchObject({
      code: 'ERR_CHART_NOT_FOUND'
    });
    await workbook.close();
  });

  it('rejects operations after close', async () => {
    const workbook = await openWorkbook({ path: fixturePath });

    await workbook.close();

    await expect(workbook.listCharts()).rejects.toBeInstanceOf(UnsupportedOperationError);
    await expect(workbook.listCharts()).rejects.toMatchObject({
      code: 'ERR_WORKBOOK_CLOSED'
    });
  });
});

function toStableIrSnapshot(model: Awaited<ReturnType<Awaited<ReturnType<typeof openWorkbook>>['getChartModel']>>) {
  return {
    chartTypes: model.chartTypes,
    legend: model.legend,
    groups: model.plotArea?.chartGroups.map((group) => ({
      type: group.type,
      axisIds: group.axisIds.length,
      ...(group.grouping ? { grouping: group.grouping } : {})
    })),
    axes: model.axes.map((axis) => ({
      kind: axis.kind,
      ...(axis.position ? { position: axis.position } : {})
    })),
    series: model.series.map((series) => ({
      chartType: series.chartType,
      points: series.points.length
    }))
  };
}

async function createDamagedRelationshipWorkbook(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B2"/>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`);
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="https://example.invalid/drawing.xml" TargetMode="External"/>
  <Relationship Id="rIdMissing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/missingDrawing.xml"/>
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`);
  zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Broken Chart"/></xdr:nvGraphicFramePr>
      <a:graphic><a:graphicData><c:chart r:id="rIdMissingChart"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
  </xdr:twoCellAnchor>
</xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMissingChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/missingChart.xml"/>
</Relationships>`);

  return zip.generateAsync({ type: 'nodebuffer' });
}

async function createManualLayoutWorkbook(): Promise<Buffer> {
  const zip = new JSZip();
  addMinimalWorkbookPackage(zip, manualLayoutChartXml());
  return zip.generateAsync({ type: 'nodebuffer' });
}

function addMinimalWorkbookPackage(zip: JSZip, chartXml: string): void {
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <drawing r:id="rIdDrawing"/>
</worksheet>`);
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`);
  zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Manual Layout Chart"/></xdr:nvGraphicFramePr>
      <a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
  </xdr:twoCellAnchor>
</xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`);
  zip.file('xl/charts/chart1.xml', chartXml);
}

function manualLayoutChartXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:layout>
        <c:manualLayout>
          <c:layoutTarget val="inner"/>
          <c:xMode val="factor"/>
          <c:yMode val="factor"/>
          <c:wMode val="factor"/>
          <c:hMode val="factor"/>
          <c:x val="0.08"/>
          <c:y val="0.12"/>
          <c:w val="0.58"/>
          <c:h val="0.76"/>
        </c:manualLayout>
      </c:layout>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>12</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="0"/>
        <c:axId val="1"/>
      </c:barChart>
      <c:catAx><c:axId val="0"/><c:axPos val="b"/><c:crossAx val="1"/></c:catAx>
      <c:valAx><c:axId val="1"/><c:axPos val="l"/><c:crossAx val="0"/></c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout>
        <c:manualLayout>
          <c:layoutTarget val="inner"/>
          <c:xMode val="factor"/>
          <c:yMode val="factor"/>
          <c:wMode val="factor"/>
          <c:hMode val="factor"/>
          <c:x val="0.71"/>
          <c:y val="0.25"/>
          <c:w val="0.24"/>
          <c:h val="0.5"/>
        </c:manualLayout>
      </c:layout>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
}

function createWorkbookLike(sheets: Record<string, MockSheet>): ExcelJsWorkbookLike {
  return {
    getWorksheet(name: string) {
      const sheet = sheets[name];
      return sheet ? {
        getCell(address: string) {
          return { value: sheet.cells[address.toUpperCase()] };
        },
        getRow(row: number) {
          return { hidden: sheet.hiddenRows?.has(row) ?? false };
        },
        getColumn(column: number) {
          return { hidden: sheet.hiddenColumns?.has(column) ?? false };
        }
      } : undefined;
    }
  };
}

interface MockSheet {
  readonly cells: Record<string, unknown>;
  readonly hiddenRows?: ReadonlySet<number>;
  readonly hiddenColumns?: ReadonlySet<number>;
}
