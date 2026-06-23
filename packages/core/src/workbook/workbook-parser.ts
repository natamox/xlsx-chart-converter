import type { PackageReader } from '../package/package-reader.js';
import { parseRelationships } from '../package/relationship-parser.js';
import { relationshipPartName, resolveRelationshipTarget } from '../package/relationship-resolver.js';
import { attr, child, children, parseXmlTree } from '../xml/xml-tree.js';
import type { WorkbookIndex } from './workbook-index.js';

const worksheetRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const chartsheetRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet';
const sharedStringsRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
const themeRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';

export async function parseWorkbookIndex(packageReader: PackageReader): Promise<WorkbookIndex> {
  const workbookPart = 'xl/workbook.xml';
  const workbookXml = (await packageReader.readPart(workbookPart)).toString('utf8');
  const workbook = parseXmlTree(workbookXml, workbookPart);
  const relsPart = relationshipPartName(workbookPart);
  const rels = await readRelationships(packageReader, relsPart);
  const sheetsNode = child(workbook, 'sheets');

  const sharedStrings = rels.find((rel) =>
    rel.type === sharedStringsRelType && rel.targetMode !== 'External'
  );
  const theme = rels.find((rel) =>
    rel.type === themeRelType && rel.targetMode !== 'External'
  );

  return {
    workbookPart,
    ...(sharedStrings
      ? { sharedStringsPart: resolveRelationshipTarget(workbookPart, sharedStrings.target) }
      : {}),
    ...(theme
      ? { themePart: resolveRelationshipTarget(workbookPart, theme.target) }
      : {}),
    sheets: children(sheetsNode ?? workbook, 'sheet').flatMap((sheet) => {
      const relId = attr(sheet, 'r:id');
      const name = attr(sheet, 'name');
      const id = attr(sheet, 'sheetId') ?? relId;
      if (!relId || !name || !id) {
        return [];
      }

      const relationship = rels.find((rel) => rel.id === relId);
      if (!relationship || relationship.targetMode === 'External') {
        return [];
      }

      const kind = relationship.type === chartsheetRelType ? 'chartsheet'
        : relationship.type === worksheetRelType ? 'worksheet'
          : undefined;
      if (!kind) {
        return [];
      }

      return [{
        id,
        relId,
        name,
        partName: resolveRelationshipTarget(workbookPart, relationship.target),
        kind
      }];
    })
  };
}

async function readRelationships(packageReader: PackageReader, relsPart: string) {
  if (!(await packageReader.hasPart(relsPart))) {
    return [];
  }
  return parseRelationships((await packageReader.readPart(relsPart)).toString('utf8'), relsPart);
}
