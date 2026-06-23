export interface ContentTypeIndex {
  readonly defaults: ReadonlyMap<string, string>;
  readonly overrides: ReadonlyMap<string, string>;
}

export function createEmptyContentTypeIndex(): ContentTypeIndex {
  return {
    defaults: new Map<string, string>(),
    overrides: new Map<string, string>()
  };
}
