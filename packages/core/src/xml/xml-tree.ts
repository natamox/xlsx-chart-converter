import { SaxesParser } from 'saxes';

import { XlsxChartError } from '../diagnostics/errors.js';

export interface XmlNode {
  readonly name: string;
  readonly localName: string;
  readonly attributes: ReadonlyMap<string, string>;
  readonly children: XmlNode[];
  text: string;
}

interface MutableXmlNode extends XmlNode {
  readonly children: MutableXmlNode[];
}

export function parseXmlTree(xml: string, sourceName: string): XmlNode {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) {
    throw new XlsxChartError('ERR_UNSAFE_XML', `Unsafe XML construct found in ${sourceName}.`);
  }

  const root: MutableXmlNode = createNode('#document', []);
  const stack: MutableXmlNode[] = [root];
  const parser = new SaxesParser({ xmlns: false });

  parser.on('opentag', (tag) => {
    const parent = stack.at(-1);
    if (!parent) {
      throw new XlsxChartError('ERR_XML_PARSE', `Invalid XML stack while parsing ${sourceName}.`);
    }

    const node = createNode(tag.name, Object.entries(tag.attributes));
    parent.children.push(node);
    stack.push(node);
  });

  parser.on('text', (text) => {
    const current = stack.at(-1);
    if (current && text.trim().length > 0) {
      current.text += text;
    }
  });

  parser.on('closetag', () => {
    stack.pop();
  });

  let parseError: unknown;
  parser.on('error', (error) => {
    parseError = error;
  });

  parser.write(xml).close();

  if (parseError) {
    throw new XlsxChartError('ERR_XML_PARSE', `Failed to parse XML part ${sourceName}.`, parseError);
  }

  const documentElement = root.children[0];
  if (!documentElement) {
    throw new XlsxChartError('ERR_XML_PARSE', `XML part ${sourceName} is empty.`);
  }

  return documentElement;
}

export function child(node: XmlNode, localName: string): XmlNode | undefined {
  return node.children.find((item) => item.localName === localName);
}

export function children(node: XmlNode, localName: string): XmlNode[] {
  return node.children.filter((item) => item.localName === localName);
}

export function childrenAny(node: XmlNode, localNames: readonly string[]): XmlNode[] {
  const wanted = new Set(localNames);
  return node.children.filter((item) => wanted.has(item.localName));
}

export function descendants(node: XmlNode, localName: string): XmlNode[] {
  const matches: XmlNode[] = [];
  walk(node, (item) => {
    if (item.localName === localName) {
      matches.push(item);
    }
  });
  return matches;
}

export function firstDescendant(node: XmlNode, localName: string): XmlNode | undefined {
  for (const item of node.children) {
    if (item.localName === localName) {
      return item;
    }
    const nested = firstDescendant(item, localName);
    if (nested) {
      return nested;
    }
  }
  return undefined;
}

export function attr(node: XmlNode | undefined, name: string): string | undefined {
  return node?.attributes.get(name);
}

export function valAttr(node: XmlNode | undefined): string | undefined {
  return attr(node, 'val');
}

export function textContent(node: XmlNode | undefined): string | undefined {
  if (!node) {
    return undefined;
  }
  const pieces: string[] = [];
  walk(node, (item) => {
    if (item.text.trim().length > 0) {
      pieces.push(item.text.trim());
    }
  });
  const text = pieces.join('');
  return text.length > 0 ? text : undefined;
}

function createNode(name: string, attrs: [string, string][]): MutableXmlNode {
  return {
    name,
    localName: name.includes(':') ? name.slice(name.indexOf(':') + 1) : name,
    attributes: new Map(attrs),
    children: [],
    text: ''
  };
}

function walk(node: XmlNode, visit: (node: XmlNode) => void): void {
  visit(node);
  for (const item of node.children) {
    walk(item, visit);
  }
}
