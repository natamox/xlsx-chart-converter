import { attr, firstDescendant, parseXmlTree } from '../xml/xml-tree.js';
import { normalizeHexColor } from './color-resolver.js';

import type { XmlNode } from '../xml/xml-tree.js';

export interface ThemeModel {
  readonly name?: string;
  readonly colors: ReadonlyMap<string, string>;
  readonly fonts: {
    readonly majorLatin?: string;
    readonly minorLatin?: string;
  };
}

export function parseThemePart(xml: string, sourceName: string): ThemeModel {
  const theme = parseXmlTree(xml, sourceName);
  const colors = new Map<string, string>();
  const colorScheme = firstDescendant(theme, 'clrScheme');
  for (const item of colorScheme?.children ?? []) {
    const color = parseThemeColor(item);
    if (color) {
      colors.set(item.localName, color);
    }
  }

  const majorLatin = attr(firstDescendant(firstDescendant(theme, 'majorFont') ?? theme, 'latin'), 'typeface');
  const minorLatin = attr(firstDescendant(firstDescendant(theme, 'minorFont') ?? theme, 'latin'), 'typeface');
  const name = attr(theme, 'name');
  return {
    ...(name ? { name } : {}),
    colors,
    fonts: {
      ...(majorLatin ? { majorLatin } : {}),
      ...(minorLatin ? { minorLatin } : {})
    }
  };
}

function parseThemeColor(node: XmlNode): string | undefined {
  const srgb = firstDescendant(node, 'srgbClr');
  const sys = firstDescendant(node, 'sysClr');
  const value = attr(srgb, 'val') ?? attr(sys, 'lastClr') ?? attr(sys, 'val');
  return value ? normalizeHexColor(value) : undefined;
}
