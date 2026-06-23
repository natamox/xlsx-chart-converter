import { UnsupportedOperationError } from '../diagnostics/errors.js';

import type { PackageReaderFactory } from './package-reader.js';

export const yauzlPackageReaderFactory: PackageReaderFactory = {
  open(): Promise<never> {
    return Promise.reject(new UnsupportedOperationError(
      'ERR_NOT_IMPLEMENTED',
      'Yauzl PackageReader is planned for M1 and is not implemented yet.'
    ));
  }
};
