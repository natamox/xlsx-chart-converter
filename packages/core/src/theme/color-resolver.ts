export function normalizeHexColor(color: string): string {
  return color.startsWith('#') ? color : `#${color}`;
}
