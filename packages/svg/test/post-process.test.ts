import { describe, expect, it } from 'vitest';
import { addSvgAccessibility, prefixSvgIds, sanitizeSvg } from '../src/index.js';

describe('SVG post-processing', () => {
  it('removes active content and external URL references', () => {
    const svg = sanitizeSvg(`
      <svg><script>alert(1)</script><foreignObject><p>bad</p></foreignObject>
      <rect onclick="evil()" href="https://example.invalid/a.svg" xlink:href='http://example.invalid/b.svg'/></svg>
    `);

    expect(svg).not.toContain('<script');
    expect(svg).not.toContain('<foreignObject');
    expect(svg).not.toContain('onclick');
    expect(svg).not.toContain('https://example.invalid');
    expect(svg).not.toContain('http://example.invalid');
  });

  it('prefixes ids and updates local references', () => {
    const svg = prefixSvgIds(
      '<svg><defs><clipPath id="clip"><path id="path"/></clipPath></defs><g clip-path="url(#clip)"><use href="#path"/></g></svg>',
      'xc-test-'
    );

    expect(svg).toContain('id="xc-test-clip"');
    expect(svg).toContain('id="xc-test-path"');
    expect(svg).toContain('url(#xc-test-clip)');
    expect(svg).toContain('href="#xc-test-path"');
  });

  it('does not double-prefix ids that already have the requested prefix', () => {
    const svg = prefixSvgIds(
      '<svg><clipPath id="xc-test-clip"/><g clip-path="url(#xc-test-clip)"/></svg>',
      'xc-test-'
    );

    expect(svg).toContain('id="xc-test-clip"');
    expect(svg).toContain('url(#xc-test-clip)');
    expect(svg).not.toContain('xc-test-xc-test-clip');
  });

  it('adds accessible title and role without duplicating existing title', () => {
    const svg = addSvgAccessibility('<svg width="100" height="80"><g/></svg>', 'Sales <Plan>');

    expect(svg).toContain('role="img"');
    expect(svg).toContain('aria-label="Sales &lt;Plan&gt;"');
    expect(svg).toContain('<title>Sales &lt;Plan&gt;</title>');
    expect(addSvgAccessibility(svg, 'Sales <Plan>').match(/<title>/g)).toHaveLength(1);
  });
});
