import type { XmlNode } from '../xml/xml-tree.js';

export interface AlternateContentChoice {
  readonly requires?: string;
  readonly xml: string;
}

const supportedRequires = new Set(['xdr', 'a', 'c', 'r']);

export function selectAlternateContent(node: XmlNode): XmlNode {
  return rewriteAlternateContent(node);
}

function rewriteAlternateContent(node: XmlNode): XmlNode {
  if (node.localName === 'AlternateContent') {
    const choice = node.children.find((child) =>
      child.localName === 'Choice' && isSupportedChoice(child.attributes.get('Requires'))
    );
    const fallback = node.children.find((child) => child.localName === 'Fallback');
    const selected = choice ?? fallback;
    if (!selected) {
      return copyNode(node, []);
    }

    const onlyChild = selected.children[0];
    if (selected.children.length === 1 && onlyChild) {
      return rewriteAlternateContent(onlyChild);
    }

    return copyNode(selected, selected.children.map(rewriteAlternateContent));
  }

  return copyNode(node, node.children.map(rewriteAlternateContent));
}

function isSupportedChoice(requires: string | undefined): boolean {
  if (!requires) {
    return true;
  }
  return requires.split(/\s+/).every((item) => supportedRequires.has(item));
}

function copyNode(node: XmlNode, children: XmlNode[]): XmlNode {
  return {
    name: node.name,
    localName: node.localName,
    attributes: node.attributes,
    children,
    text: node.text
  };
}
