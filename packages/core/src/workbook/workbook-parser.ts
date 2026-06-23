import type { PackageReader } from '../package/package-reader.js';
import type { WorkbookIndex } from './workbook-index.js';

export function parseWorkbookIndex(packageReader: PackageReader): Promise<WorkbookIndex> {
  void packageReader;
  return Promise.resolve({
    workbookPart: '/xl/workbook.xml',
    sheets: []
  });
}
