export class XlsxChartError extends Error {
  override readonly name: string = 'XlsxChartError';

  constructor(
    readonly code: string,
    message: string,
    readonly cause?: unknown
  ) {
    super(message);
  }
}

export class UnsupportedOperationError extends XlsxChartError {
  override readonly name: string = 'UnsupportedOperationError';
}

export function isXlsxChartError(error: unknown): error is XlsxChartError {
  return error instanceof XlsxChartError;
}
