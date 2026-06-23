import type { AnchorModel } from './anchor-model.js';

export function resolveAnchorSize(anchor: AnchorModel): { width: number; height: number } {
  return {
    width: anchor.widthPx ?? 640,
    height: anchor.heightPx ?? 360
  };
}
