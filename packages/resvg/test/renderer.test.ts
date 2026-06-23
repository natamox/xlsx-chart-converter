import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

import { describe, expect, it } from 'vitest';
import { ResvgPngRenderer } from '../src/index.js';

describe('ResvgPngRenderer', () => {
  it('renders PNG with scale and explicit background', async () => {
    const renderer = new ResvgPngRenderer();
    const png = await renderer.render(
      '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="8"><rect width="10" height="8" fill="red"/></svg>',
      { width: 10, height: 8, scale: 2, background: '#ffffff' }
    );

    expect(png.subarray(0, 8).toString('hex')).toBe('89504e470d0a1a0a');
    expect(readPngSize(png)).toEqual({ width: 20, height: 16 });
  });

  it('reports missing custom font files without failing rendering', async () => {
    const renderer = new ResvgPngRenderer();
    const missingFont = path.join(os.tmpdir(), 'missing-excel-chart-font.ttf');

    const result = await renderer.renderWithDiagnostics(
      '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="8"><text x="1" y="7">字体</text></svg>',
      { width: 10, height: 8, scale: 1, fonts: { fontFiles: [missingFont] } }
    );

    expect(result.data.subarray(0, 8).toString('hex')).toBe('89504e470d0a1a0a');
    expect(result.diagnostics).toMatchObject([{
      code: 'FONT_FILE_NOT_FOUND',
      severity: 'warning',
      path: missingFont
    }]);

    await fs.rm(missingFont, { force: true });
  });
});

function readPngSize(png: Buffer): { width: number; height: number } {
  return {
    width: png.readUInt32BE(16),
    height: png.readUInt32BE(20)
  };
}
