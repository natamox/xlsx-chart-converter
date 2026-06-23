import type { Diagnostic } from '../public/types.js';

export class DiagnosticCollector {
  private readonly diagnostics: Diagnostic[] = [];

  add(diagnostic: Diagnostic): void {
    this.diagnostics.push(diagnostic);
  }

  list(): readonly Diagnostic[] {
    return this.diagnostics;
  }
}
