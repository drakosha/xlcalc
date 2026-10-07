/**
 * Colours of format sections: the eight names and `[ColorN]`, an index into Excel's default palette of 56 colours
 * (Microsoft, "Review guidelines for customizing a number format"; the palette is xlsx's indexed colours 8–63).
 */

/** The default palette, `[Color1]` to `[Color56]`. */
export const PALETTE: readonly string[] = (
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 ' +
  '0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 ' +
  'FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333'
)
  .split(' ')
  .map((c) => `#${c}`)

/** The eight colour names in the order black, white, red, green, blue, yellow, magenta, cyan: `[Color1]` to `[Color8]`. */
export const colorNames = (names: readonly string[]): ReadonlyMap<string, string> => new Map(names.map((n, i) => [n.toLowerCase(), PALETTE[i]]))

/** en-US names (and the names an xlsx file stores). */
export const EN_COLORS = colorNames(['Black', 'White', 'Red', 'Green', 'Blue', 'Yellow', 'Magenta', 'Cyan'])
