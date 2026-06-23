import { UnsupportedOperationError } from '../../diagnostics/errors.js';

export function parseClassicChartPart(): Promise<never> {
  return Promise.reject(new UnsupportedOperationError(
    'ERR_NOT_IMPLEMENTED',
    'Classic chart parsing is planned for M2 and is not implemented yet.'
  ));
}
