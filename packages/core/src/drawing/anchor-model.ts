export interface AnchorModel {
  readonly kind: 'twoCell' | 'oneCell' | 'absolute';
  readonly widthPx?: number;
  readonly heightPx?: number;
}
