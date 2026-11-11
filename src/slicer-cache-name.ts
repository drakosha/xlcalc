/** Whether a defined name's formula is `#N/A`, the value Excel gives the global name of a slicer or timeline cache. */
export const isCacheNameValue = (ref: string): boolean => /^=?\s*#N\/A\s*$/i.test(ref)
