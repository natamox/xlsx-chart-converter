import type { Diagnostic } from '../public/types.js';

export function sortDiagnostics(diagnostics: readonly Diagnostic[]): Diagnostic[] {
  return [...diagnostics].sort((left, right) => left.code.localeCompare(right.code));
}
