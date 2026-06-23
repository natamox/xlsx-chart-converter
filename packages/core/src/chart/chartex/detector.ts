export function isChartExPart(contentType: string | undefined): boolean {
  return contentType?.includes('chartEx') ?? false;
}
