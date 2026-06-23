export interface ParsedA1Reference {
  readonly sheetName?: string;
  readonly start: CellAddress;
  readonly end: CellAddress;
}

export interface CellAddress {
  readonly col: number;
  readonly row: number;
}

export function parseA1Reference(formula: string, defaultSheetName?: string): ParsedA1Reference | undefined {
  const normalized = formula.trim();
  const match = /^(?:(?<sheet>'(?:[^']|'')+'|[^!]+)!)?(?<start>\$?[A-Za-z]{1,3}\$?\d+)(?::(?<end>\$?[A-Za-z]{1,3}\$?\d+))?$/.exec(normalized);
  if (!match?.groups?.start) {
    return undefined;
  }

  const sheetName = match.groups.sheet ? unquoteSheetName(match.groups.sheet) : defaultSheetName;
  if (!sheetName) {
    return undefined;
  }

  const start = parseAddress(match.groups.start);
  const end = parseAddress(match.groups.end ?? match.groups.start);
  if (!start || !end) {
    return undefined;
  }

  return {
    sheetName,
    start: {
      col: Math.min(start.col, end.col),
      row: Math.min(start.row, end.row)
    },
    end: {
      col: Math.max(start.col, end.col),
      row: Math.max(start.row, end.row)
    }
  };
}

export function addressKey(address: CellAddress): string {
  return `${columnNumberToName(address.col)}${address.row}`;
}

export function columnNumberToName(column: number): string {
  let value = column;
  let name = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function unquoteSheetName(sheetName: string): string {
  if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
    return sheetName.slice(1, -1).replaceAll("''", "'");
  }
  return sheetName;
}

function parseAddress(value: string): CellAddress | undefined {
  const match = /^\$?(?<col>[A-Za-z]{1,3})\$?(?<row>\d+)$/.exec(value);
  if (!match?.groups?.col || !match.groups.row) {
    return undefined;
  }

  return {
    col: columnNameToNumber(match.groups.col),
    row: Number(match.groups.row)
  };
}

function columnNameToNumber(name: string): number {
  let value = 0;
  for (const char of name.toUpperCase()) {
    value = value * 26 + char.charCodeAt(0) - 64;
  }
  return value;
}
