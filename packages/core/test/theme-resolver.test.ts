import { describe, expect, it } from 'vitest';
import { parseThemePart } from '../src/theme/theme-parser.js';
import { resolveChartStyle, resolveChartStyleTree } from '../src/theme/style-resolver.js';
import { normalizeHexColor, resolveSchemeColor } from '../src/theme/color-resolver.js';
import { parseXmlTree } from '../src/xml/xml-tree.js';

describe('theme and style resolution', () => {
  it('parses Office theme colors and fonts', () => {
    const theme = parseThemePart(`<?xml version="1.0" encoding="UTF-8"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom">
        <a:themeElements>
          <a:clrScheme name="Office">
            <a:dk1><a:sysClr lastClr="000000"/></a:dk1>
            <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
          </a:clrScheme>
          <a:fontScheme name="Fonts">
            <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
            <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
          </a:fontScheme>
        </a:themeElements>
      </a:theme>`, 'xl/theme/theme1.xml');

    expect(theme.name).toBe('Custom');
    expect(theme.colors.get('accent1')).toBe('#4472C4');
    expect(theme.fonts.majorLatin).toBe('Aptos Display');
    expect(theme.fonts.minorLatin).toBe('Aptos');
  });

  it('resolves solid fills, lines, alpha, and theme scheme colors', () => {
    const theme = parseThemePart(`<?xml version="1.0" encoding="UTF-8"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements><a:clrScheme name="Office">
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        </a:clrScheme></a:themeElements>
      </a:theme>`, 'xl/theme/theme1.xml');

    const style = resolveChartStyle(`<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:spPr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill></c:spPr>
        <c:chart>
          <c:plotArea>
            <c:spPr><a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill></c:spPr>
            <c:barChart><c:ser>
              <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
                <a:ln w="19050"><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:ln>
              </c:spPr>
            </c:ser></c:barChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`, theme);

    expect(normalizeHexColor('abc')).toBe('#AABBCC');
    expect(resolveSchemeColor('accent1', theme)).toBe('#4472C4');
    expect(style.chartArea?.fill?.color).toBe('#F2F2F2');
    expect(style.plotArea?.fill).toMatchObject({ kind: 'solid', color: '#4472C4', alpha: 0.5 });
    expect(style.series).toBeDefined();
    const [firstSeries] = style.series ?? [];
    expect(firstSeries?.fill?.color).toBe('#4472C4');
    expect(firstSeries?.line).toEqual({ color: '#111111', width: 1.5 });
  });

  it('ignores empty line style and applies basic luminance transforms', () => {
    const theme = parseThemePart(`<?xml version="1.0" encoding="UTF-8"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements><a:clrScheme name="Office">
          <a:accent1><a:srgbClr val="808080"/></a:accent1>
        </a:clrScheme></a:themeElements>
      </a:theme>`, 'xl/theme/theme1.xml');

    const style = resolveChartStyle(`<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart><c:plotArea><c:barChart><c:ser>
          <c:spPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill><a:ln/></c:spPr>
        </c:ser></c:barChart></c:plotArea></c:chart>
      </c:chartSpace>`, theme);

    const [firstSeries] = style.series ?? [];
    expect(firstSeries?.fill?.color).toBe('#808080');
    expect(firstSeries?.fill?.transformedColor).toBe('#404040');
    expect(firstSeries?.line).toBeUndefined();
  });

  it('resolves F1 style cascade, text, marker, per-point style, and unsupported fills', () => {
    const theme = parseThemePart(`<?xml version="1.0" encoding="UTF-8"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements><a:clrScheme name="Office">
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
          <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
        </a:clrScheme></a:themeElements>
      </a:theme>`, 'xl/theme/theme1.xml');

    const style = resolveChartStyle(`<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart>
          <c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:latin typeface="Aptos Display"/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>
          <c:legend><c:txPr><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:legend>
          <c:plotArea>
            <c:catAx><c:axId val="cat"/><c:txPr><a:p><a:pPr><a:defRPr i="1"><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:catAx>
            <c:barChart><c:ser>
              <c:spPr><a:noFill/><a:ln w="25400"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
              <c:marker><c:symbol val="diamond"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></c:spPr></c:marker>
              <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>
              <c:dPt><c:idx val="2"/><c:spPr><a:gradFill/></c:spPr></c:dPt>
            </c:ser></c:barChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`, theme, {
      chartColorStyleXml: `<?xml version="1.0" encoding="UTF-8"?>
        <cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <cs:variation><a:schemeClr val="accent1"/></cs:variation>
        </cs:colorStyle>`
    });

    expect(style.title).toMatchObject({ fontFamily: 'Aptos Display', fontSize: 14, bold: true, color: '#222222' });
    expect(style.legend).toMatchObject({ fontSize: 9, color: '#333333' });
    expect(style.axes?.[0]).toMatchObject({ axisId: 'cat', text: { italic: true, color: '#444444' } });
    expect(style.series?.[0]).toMatchObject({
      fill: { kind: 'none' },
      line: { color: '#ED7D31', width: 2, dash: 'dash' }
    });
    expect(style.seriesMarkers?.[0]).toMatchObject({ symbol: 'diamond', size: 7, fill: { color: '#00AA00' } });
    expect(style.pointStyles?.[0]?.[0]).toMatchObject({ index: 1, style: { fill: { color: '#FF0000' } } });
    expect(style.pointStyles?.[0]?.[1]).toMatchObject({ index: 2, style: { fill: { kind: 'gradient' } } });
  });

  it('returns diagnostics for unsupported F1 fill styles', () => {
    const chartSpace = parseXmlTree(`<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart><c:plotArea><c:barChart><c:ser>
          <c:spPr><a:blipFill/></c:spPr>
        </c:ser></c:barChart></c:plotArea></c:chart>
      </c:chartSpace>`, 'chart.xml');

    const result = resolveChartStyleTree(chartSpace, undefined);

    expect(result.style.series?.[0]?.fill?.kind).toBe('picture');
    expect(result.diagnostics).toContainEqual(expect.objectContaining({
      code: 'UNSUPPORTED_FILL_STYLE',
      severity: 'warning',
      details: { kind: 'picture' }
    }));
  });
});
