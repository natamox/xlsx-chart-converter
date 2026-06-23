import { attr, child, firstDescendant } from '../xml/xml-tree.js';

import type { ThemeModel } from './theme-parser.js';
import type { XmlNode } from '../xml/xml-tree.js';

export interface ResolvedColor {
  readonly color: string;
  readonly transformedColor?: string;
  readonly alpha?: number;
}

export function normalizeHexColor(color: string): string {
  const normalized = color.trim().replace(/^#/, '');
  const expanded = normalized.length === 3
    ? normalized.split('').map((item) => `${item}${item}`).join('')
    : normalized;
  return `#${expanded.toUpperCase()}`;
}

export function resolveSchemeColor(name: string, theme: ThemeModel | undefined): string | undefined {
  return theme?.colors.get(name);
}

export function resolveDrawingColor(node: XmlNode | undefined, theme: ThemeModel | undefined): ResolvedColor | undefined {
  const srgb = firstDescendant(node ?? emptyNode, 'srgbClr');
  const scheme = firstDescendant(node ?? emptyNode, 'schemeClr');
  const value = attr(srgb, 'val') ?? resolveSchemeColor(attr(scheme, 'val') ?? '', theme);
  if (!value) {
    return undefined;
  }

  const alphaNode = child(srgb ?? scheme ?? emptyNode, 'alpha');
  const alpha = parseAlpha(attr(alphaNode, 'val'));
  const color = normalizeHexColor(value);
  const transformedColor = applyColorTransforms(color, srgb ?? scheme);
  return {
    color,
    ...(transformedColor === color ? {} : { transformedColor }),
    ...(alpha === undefined ? {} : { alpha })
  };
}

function applyColorTransforms(color: string, node: XmlNode | undefined): string {
  const rgb = hexToRgb(color);
  if (!rgb || !node) {
    return color;
  }

  const lumMod = parsePercentage(attr(child(node, 'lumMod'), 'val'));
  const lumOff = parsePercentage(attr(child(node, 'lumOff'), 'val'));
  if (lumMod === undefined && lumOff === undefined) {
    return color;
  }

  return rgbToHex({
    r: clamp(rgb.r * (lumMod ?? 1) + 255 * (lumOff ?? 0)),
    g: clamp(rgb.g * (lumMod ?? 1) + 255 * (lumOff ?? 0)),
    b: clamp(rgb.b * (lumMod ?? 1) + 255 * (lumOff ?? 0))
  });
}

function parsePercentage(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / 100000 : undefined;
}

function hexToRgb(color: string): { r: number; g: number; b: number } | undefined {
  const match = /^#([0-9A-F]{6})$/i.exec(color);
  const value = match?.[1];
  if (!value) {
    return undefined;
  }
  return {
    r: Number.parseInt(value.slice(0, 2), 16),
    g: Number.parseInt(value.slice(2, 4), 16),
    b: Number.parseInt(value.slice(4, 6), 16)
  };
}

function rgbToHex(rgb: { r: number; g: number; b: number }): string {
  return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`;
}

function toHex(value: number): string {
  return value.toString(16).padStart(2, '0').toUpperCase();
}

function clamp(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function parseAlpha(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) ? Math.max(0, Math.min(1, numeric / 100000)) : undefined;
}

const emptyNode: XmlNode = {
  name: '#empty',
  localName: '#empty',
  attributes: new Map(),
  children: [],
  text: ''
};
