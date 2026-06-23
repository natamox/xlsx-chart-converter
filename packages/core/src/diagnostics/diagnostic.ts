import type { Diagnostic, DiagnosticSeverity } from '../public/types.js';

export function createDiagnostic(
  code: string,
  severity: DiagnosticSeverity,
  message: string,
  details?: Record<string, unknown>
): Diagnostic {
  return {
    code,
    severity,
    message,
    ...(details === undefined ? {} : { details })
  };
}
