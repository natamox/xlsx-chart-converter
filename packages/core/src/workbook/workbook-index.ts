export interface WorkbookSheetIndex {
  readonly id: string;
  readonly name: string;
  readonly partName: string;
  readonly kind: 'worksheet' | 'chartsheet';
}

export interface WorkbookIndex {
  readonly workbookPart: string;
  readonly sheets: readonly WorkbookSheetIndex[];
}
