export function sanitizeSvg(svg: string): string {
  return svg
    .replace(/<!DOCTYPE\b[\s\S]*?>/gi, '')
    .replace(/<script\b[\s\S]*?<\/script>/gi, '')
    .replace(/<foreignObject\b[\s\S]*?<\/foreignObject>/gi, '')
    .replace(/\s+on[a-z]+\s*=\s*"[^"]*"/gi, '')
    .replace(/\s+on[a-z]+\s*=\s*'[^']*'/gi, '')
    .replace(/\s(?:href|xlink:href)\s*=\s*"(?:https?:|file:|javascript:)[^"]*"/gi, '')
    .replace(/\s(?:href|xlink:href)\s*=\s*'(?:https?:|file:|javascript:)[^']*'/gi, '')
    .replace(/url\((['"]?)(?:https?:|file:|javascript:)[^)]+\1\)/gi, 'none');
}
