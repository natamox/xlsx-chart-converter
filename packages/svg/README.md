# @natamox/excel-chart-svg

SVG post-processing helpers for excel-chart renderer output.

This package owns SVG sanitizing, ID prefixing, and accessibility metadata helpers. It stays separate from renderers so SVG safety and normalization can be reused.

## Usage

```ts
import { addSvgAccessibility, prefixSvgIds, sanitizeSvg } from '@natamox/excel-chart-svg';

const safeSvg = addSvgAccessibility(
  prefixSvgIds(sanitizeSvg(svg), 'chart-'),
  'Revenue by quarter'
);
```

## Helpers

- `sanitizeSvg(svg)`: removes doctype, scripts, foreign objects, event attributes, unsafe external hrefs, and unsafe external URL references.
- `prefixSvgIds(svg, prefix)`: prefixes `id` attributes and matching local `url(#id)` / `#id` references.
- `addSvgAccessibility(svg, title)`: adds `role="img"`, `aria-label`, and a normalized `<title>`.

## Limits

- The sanitizer is a deterministic string-based post-processor for renderer output, not a general-purpose untrusted HTML/XML sanitizer.
- Remote, file, and JavaScript references are stripped; local fragment references are preserved and can be prefixed.

## Runtime

Node.js >= 22.
