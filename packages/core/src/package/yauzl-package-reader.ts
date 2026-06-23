import fs from 'node:fs/promises';
import { Readable } from 'node:stream';

import yauzl from 'yauzl';

import { XlsxChartError } from '../diagnostics/errors.js';
import { defaultPackageLimits } from './package-limits.js';
import { normalizePartName } from './relationship-resolver.js';

import type { PackageReaderFactory } from './package-reader.js';
import type { PackagePart, PackageReader } from './package-reader.js';

type ZipFile = yauzl.ZipFile;
type Entry = yauzl.Entry;

class YauzlPackagePart implements PackagePart {
  constructor(
    readonly name: string,
    private readonly reader: YauzlPackageReader
  ) {}

  read(): Promise<Buffer> {
    return this.reader.readPart(this.name);
  }
}

class YauzlPackageReader implements PackageReader {
  private readonly entries = new Map<string, Entry>();
  private readonly normalizedNames = new Map<string, string>();
  private parts: PackagePart[] = [];

  constructor(private readonly zipFile: ZipFile) {}

  index(entry: Entry): void {
    if (!entry.fileName.endsWith('/')) {
      const name = normalizePartName(entry.fileName);
      this.entries.set(name, entry);
      this.normalizedNames.set(name.toLowerCase(), name);
      this.parts = [...this.entries.keys()].sort().map((partName) =>
        new YauzlPackagePart(partName, this)
      );
    }
  }

  listParts(): Promise<readonly PackagePart[]> {
    return Promise.resolve(this.parts);
  }

  async readPart(partName: string): Promise<Buffer> {
    const name = await this.requireResolvedPartName(partName);
    const entry = this.entries.get(name);
    if (!entry) {
      throw new XlsxChartError('ERR_PACKAGE_PART_NOT_FOUND', `Package part not found: ${name}.`);
    }

    if (entry.uncompressedSize > defaultPackageLimits.maxEntryBytes) {
      throw new XlsxChartError('ERR_PACKAGE_ENTRY_TOO_LARGE', `Package part is too large: ${name}.`);
    }

    return readEntry(this.zipFile, entry);
  }

  hasPart(partName: string): Promise<boolean> {
    return Promise.resolve(this.resolvePartNameSync(partName) !== undefined);
  }

  resolvePartName(partName: string): Promise<string | undefined> {
    return Promise.resolve(this.resolvePartNameSync(partName));
  }

  close(): Promise<void> {
    this.zipFile.close();
    return Promise.resolve();
  }

  private async requireResolvedPartName(partName: string): Promise<string> {
    const resolved = await this.resolvePartName(partName);
    if (!resolved) {
      throw new XlsxChartError(
        'ERR_PACKAGE_PART_NOT_FOUND',
        `Package part not found: ${normalizePartName(partName)}.`
      );
    }
    return resolved;
  }

  private resolvePartNameSync(partName: string): string | undefined {
    const normalized = normalizePartName(partName);
    return this.entries.has(normalized)
      ? normalized
      : this.normalizedNames.get(normalized.toLowerCase());
  }
}

export const yauzlPackageReaderFactory: PackageReaderFactory = {
  async open(source): Promise<PackageReader> {
    const buffer = await normalizeSource(source);
    if (buffer.length > defaultPackageLimits.maxInputBytes) {
      throw new XlsxChartError('ERR_INPUT_TOO_LARGE', 'Workbook exceeds the configured size limit.');
    }

    return openReader(buffer);
  }
};

async function normalizeSource(source: Buffer | Uint8Array | { path: string }): Promise<Buffer> {
  if (Buffer.isBuffer(source)) {
    return source;
  }
  if (source instanceof Uint8Array) {
    return Buffer.from(source);
  }
  return fs.readFile(source.path);
}

function openReader(buffer: Buffer): Promise<PackageReader> {
  return new Promise((resolve, reject) => {
    yauzl.fromBuffer(buffer, { lazyEntries: true, validateEntrySizes: true }, (error, zipFile) => {
      if (error) {
        reject(new XlsxChartError('ERR_INVALID_PACKAGE', 'Failed to open OOXML ZIP package.', error));
        return;
      }
      if (!zipFile) {
        reject(new XlsxChartError('ERR_INVALID_PACKAGE', 'Failed to open OOXML ZIP package.'));
        return;
      }
      if (zipFile.entryCount > defaultPackageLimits.maxEntries) {
        zipFile.close();
        reject(new XlsxChartError('ERR_TOO_MANY_PACKAGE_ENTRIES', 'Workbook has too many ZIP entries.'));
        return;
      }

      const reader = new YauzlPackageReader(zipFile);
      zipFile.on('entry', (entry: Entry) => {
        reader.index(entry);
        zipFile.readEntry();
      });
      zipFile.on('error', (zipError) => {
        reject(new XlsxChartError('ERR_INVALID_PACKAGE', 'Failed to index OOXML ZIP package.', zipError));
      });
      zipFile.on('end', () => {
        resolve(reader);
      });
      zipFile.readEntry();
    });
  });
}

function readEntry(zipFile: ZipFile, entry: Entry): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    zipFile.openReadStream(entry, (error, stream) => {
      if (error) {
        reject(new XlsxChartError('ERR_PACKAGE_READ_FAILED', `Failed to read ${entry.fileName}.`, error));
        return;
      }
      if (!stream) {
        reject(new XlsxChartError('ERR_PACKAGE_READ_FAILED', `Failed to read ${entry.fileName}.`));
        return;
      }

      const chunks: Buffer[] = [];
      Readable.from(stream)
        .on('data', (chunk: Buffer) => {
          chunks.push(chunk);
        })
        .on('error', (streamError) => {
          reject(new XlsxChartError(
            'ERR_PACKAGE_READ_FAILED',
            `Failed to read ${entry.fileName}.`,
            streamError
          ));
        })
        .on('end', () => {
          resolve(Buffer.concat(chunks));
        });
    });
  });
}
