/** Normal fonts in the geometry cases, with the three point sizes measured on both platforms. */
export const NORMAL_FONTS = ['Arial', 'Arial Cyr', 'Times New Roman'].flatMap((name) =>
  [8, 10, 11].map((size) => ({ id: `${name.toLowerCase().replaceAll(' ', '-')}-${size}`, bookFont: { name, size } })),
)
