export function addSvgAccessibility(svg: string, title: string | undefined): string {
  const accessibleTitle = escapeXml(title ?? 'Chart');
  const openTagStart = svg.search(/<svg\b/i);
  if (openTagStart < 0) {
    return svg;
  }

  const openTagEnd = findTagEnd(svg, openTagStart);
  if (openTagEnd < 0) {
    return svg;
  }

  let openTag = svg.slice(openTagStart, openTagEnd + 1);
  if (!/\srole=/.test(openTag)) {
    openTag = openTag.replace(/>$/, ' role="img">');
  }
  if (!/\saria-label=/.test(openTag)) {
    openTag = openTag.replace(/>$/, ` aria-label="${accessibleTitle}">`);
  }

  const withoutExistingTitle = svg.replace(/<title>[\s\S]*?<\/title>/i, '');
  const adjustedOpenTagEnd = findTagEnd(withoutExistingTitle, openTagStart);
  if (adjustedOpenTagEnd < 0) {
    return withoutExistingTitle;
  }
  return `${withoutExistingTitle.slice(0, openTagStart)}${openTag}<title>${accessibleTitle}</title>${withoutExistingTitle.slice(adjustedOpenTagEnd + 1)}`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function findTagEnd(svg: string, start: number): number {
  let quote: string | undefined;
  for (let index = start; index < svg.length; index += 1) {
    const character = svg[index];
    if (quote) {
      if (character === quote) {
        quote = undefined;
      }
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === '>') {
      return index;
    }
  }
  return -1;
}
