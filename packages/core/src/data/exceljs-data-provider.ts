export interface ExcelJsDataProvider {
  getCellValue(sheetName: string, address: string): unknown;
}
