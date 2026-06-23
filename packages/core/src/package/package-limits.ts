export interface PackageLimits {
  readonly maxInputBytes: number;
  readonly maxEntries: number;
  readonly maxEntryBytes: number;
  readonly maxTotalInflatedBytes: number;
  readonly maxCompressionRatio: number;
}

export const defaultPackageLimits: PackageLimits = {
  maxInputBytes: 100 * 1024 * 1024,
  maxEntries: 10_000,
  maxEntryBytes: 50 * 1024 * 1024,
  maxTotalInflatedBytes: 250 * 1024 * 1024,
  maxCompressionRatio: 100
};
