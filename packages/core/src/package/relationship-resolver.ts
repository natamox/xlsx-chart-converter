import path from 'node:path';

import { XlsxChartError } from '../diagnostics/errors.js';

export function isExternalRelationship(targetMode: string | undefined): boolean {
  return targetMode === 'External';
}

export function relationshipPartName(sourcePart: string): string {
  const normalized = normalizePartName(sourcePart);
  return path.posix.join(
    path.posix.dirname(normalized),
    '_rels',
    `${path.posix.basename(normalized)}.rels`
  );
}

export function resolveRelationshipTarget(sourcePart: string, target: string): string {
  const decoded = target.replace(/\\/g, '/');
  const base = path.posix.dirname(normalizePartName(sourcePart));
  const resolved = decoded.startsWith('/')
    ? path.posix.normalize(decoded.slice(1))
    : path.posix.normalize(path.posix.join(base, decoded));

  if (
    resolved === '..' ||
    resolved.startsWith('../') ||
    resolved.includes('/../') ||
    resolved.includes('\0')
  ) {
    throw new XlsxChartError(
      'ERR_RELATIONSHIP_PATH_ESCAPE',
      `Relationship target escapes the package: ${target}.`
    );
  }

  return resolved;
}

export function normalizePartName(partName: string): string {
  return partName.replace(/\\/g, '/').replace(/^\/+/, '');
}
