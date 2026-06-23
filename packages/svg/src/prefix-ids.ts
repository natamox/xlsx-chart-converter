export function prefixSvgIds(svg: string, prefix: string): string {
  if (!prefix) {
    return svg;
  }

  const ids = new Set<string>();
  const withIds = svg.replace(/\sid=(["'])([^"']+)\1/g, (match, quote: string, id: string) => {
    if (id.startsWith(prefix)) {
      ids.add(id);
      return match;
    }
    ids.add(id);
    return ` id=${quote}${prefix}${id}${quote}`;
  });

  let result = withIds;
  for (const id of ids) {
    if (id.startsWith(prefix)) {
      continue;
    }
    const prefixed = id.startsWith(prefix) ? id : `${prefix}${id}`;
    result = result
      .replace(new RegExp(`url\\(#${escapeRegExp(id)}\\)`, 'g'), `url(#${prefixed})`)
      .replace(new RegExp(`(["'])#${escapeRegExp(id)}\\1`, 'g'), `$1#${prefixed}$1`);
  }
  return result;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
