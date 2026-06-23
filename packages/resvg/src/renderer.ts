import fs from 'node:fs/promises';

import { Resvg } from '@resvg/resvg-js';

import type { Diagnostic, PngRenderer, RenderContext } from '@natamox/excel-chart-core';

export interface ResvgPngRendererOptions {
  readonly defaultBackground?: string;
}

export class ResvgPngRenderer implements PngRenderer {
  constructor(private readonly options: ResvgPngRendererOptions = {}) {
    void this.options;
  }

  render(svg: string, context: RenderContext): Promise<Buffer> {
    return this.renderWithDiagnostics(svg, context).then((result) => result.data);
  }

  async renderWithDiagnostics(svg: string, context: RenderContext): Promise<{
    readonly data: Buffer;
    readonly diagnostics: readonly Diagnostic[];
  }> {
    const diagnostics = await validateFontFiles(context.fonts?.fontFiles);
    const resvg = new Resvg(svg, {
      ...(context.background ?? this.options.defaultBackground
        ? { background: context.background ?? this.options.defaultBackground }
        : {}),
      fitTo: {
        mode: 'zoom',
        value: context.scale
      },
      font: {
        loadSystemFonts: true,
        ...(context.fonts?.fontFiles ? { fontFiles: [...context.fonts.fontFiles] } : {})
      }
    });

    return {
      data: resvg.render().asPng(),
      diagnostics
    };
  }
}

async function validateFontFiles(fontFiles: readonly string[] | undefined): Promise<Diagnostic[]> {
  const diagnostics: Diagnostic[] = [];
  for (const fontFile of fontFiles ?? []) {
    try {
      await fs.access(fontFile);
    } catch {
      diagnostics.push({
        code: 'FONT_FILE_NOT_FOUND',
        severity: 'warning',
        message: `Custom font file was not found: ${fontFile}`,
        path: fontFile
      });
    }
  }
  return diagnostics;
}
