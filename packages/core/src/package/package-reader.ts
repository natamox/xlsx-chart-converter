export interface PackagePart {
  readonly name: string;
  readonly contentType?: string;
  read(): Promise<Buffer>;
}

export interface PackageReader {
  listParts(): Promise<readonly PackagePart[]>;
  readPart(partName: string): Promise<Buffer>;
  hasPart(partName: string): Promise<boolean>;
  resolvePartName(partName: string): Promise<string | undefined>;
  close(): Promise<void>;
}

export interface PackageReaderFactory {
  open(source: Buffer | Uint8Array | { path: string }): Promise<PackageReader>;
}
