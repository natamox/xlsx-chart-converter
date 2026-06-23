declare module 'yauzl' {
  import type { Readable } from 'node:stream';

  export interface Entry {
    readonly fileName: string;
    readonly uncompressedSize: number;
  }

  export interface ZipFile {
    readonly entryCount: number;
    openReadStream(entry: Entry, callback: (error: Error | null, stream?: Readable) => void): void;
    readEntry(): void;
    on(event: 'entry', callback: (entry: Entry) => void): void;
    on(event: 'end', callback: () => void): void;
    on(event: 'error', callback: (error: Error) => void): void;
    close(): void;
  }

  export function fromBuffer(
    buffer: Buffer,
    options: { lazyEntries: boolean; validateEntrySizes: boolean },
    callback: (error: Error | null, zipFile?: ZipFile) => void
  ): void;
}
