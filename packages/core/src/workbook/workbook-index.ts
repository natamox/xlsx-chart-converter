export interface WorkbookSheetIndex {
  readonly id: string;
  readonly relId: string;
  readonly name: string;
  readonly partName: string;
  readonly kind: 'worksheet' | 'chartsheet';
}

export interface WorkbookIndex {
  readonly workbookPart: string;
  readonly sharedStringsPart?: string;
  readonly themePart?: string;
  readonly sheets: readonly WorkbookSheetIndex[];
}
