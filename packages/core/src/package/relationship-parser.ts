export interface PackageRelationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: 'Internal' | 'External';
}
