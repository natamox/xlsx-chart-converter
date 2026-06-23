# @natamox/excel-chart-svg

SVG post-processing helpers for excel-chart renderer output.

This package owns SVG sanitizing, ID prefixing, and accessibility metadata helpers. It stays separate from renderers so SVG safety and normalization can be reused.

## Usage

```ts
import { addSvgAccessibility, prefixSvgIds, sanitizeSvg } from '@natamox/excel-chart-svg';

const safeSvg = addSvgAccessibility(prefixSvgIds(sanitizeSvg(svg), 'chart-'), 'Chart');
```

## Status

The package is currently an M0 scaffold. Public helper names are present; sanitizer, ID rewriting, and accessibility behavior are still being implemented.

## Runtime

Node.js >= 22.
