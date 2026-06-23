import { attr, children, parseXmlTree } from '../xml/xml-tree.js';

export interface PackageRelationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: 'Internal' | 'External';
}

export function parseRelationships(xml: string, sourceName: string): PackageRelationship[] {
  const root = parseXmlTree(xml, sourceName);

  return children(root, 'Relationship').flatMap((node) => {
    const id = attr(node, 'Id');
    const type = attr(node, 'Type');
    const target = attr(node, 'Target');
    const targetMode = attr(node, 'TargetMode');

    if (!id || !type || !target) {
      return [];
    }

    return [{
      id,
      type,
      target,
      targetMode: targetMode === 'External' ? 'External' : 'Internal'
    }];
  });
}
