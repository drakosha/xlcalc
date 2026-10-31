/** The size Excel gives a new slicer and a new timeline, in EMU (`addSlicer`, `addTimeline` without a place of their own; the set slicer-objects): literals alone, a light module of the root and `xlcalc/text`. */

/** A new slicer's width and height in EMU, as Excel gives them: 2 by 2.76 inches. */
export const DEFAULT_SLICER_SIZE: Readonly<{ width: number; height: number }> = /* @__PURE__ */ Object.freeze({ width: 1_828_800, height: 2_524_125 })

/** A new timeline's width and height in EMU, as Excel gives them: 3.65 by 1.5 inches. */
export const DEFAULT_TIMELINE_SIZE: Readonly<{ width: number; height: number }> = /* @__PURE__ */ Object.freeze({ width: 3_333_750, height: 1_371_600 })
