export interface AnchorModel {
  readonly kind: 'twoCell' | 'oneCell' | 'absolute';
  readonly from?: AnchorMarker;
  readonly to?: AnchorMarker;
  readonly ext?: AnchorExtent;
}

export interface AnchorMarker {
  readonly col: number;
  readonly row: number;
  readonly colOffEmu: number;
  readonly rowOffEmu: number;
}

export interface AnchorExtent {
  readonly cxEmu: number;
  readonly cyEmu: number;
}
