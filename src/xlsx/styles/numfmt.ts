/**
 * Built-in number formats: a file stores only their number. The codes are the ones Excel for Mac en-US gives them
 * (measured: `test/excel/xlsx/formats.json`). The date and time formats (14 to 22, 27 to 36, 50 to 58) and the
 * currency ones (5 to 8, 41 to 44) Excel shows by the regional settings, so a file of another locale or platform
 * shows them otherwise (14 is `m/d/yyyy` in Excel for Windows en-US, `dd.mm.yyyy` in ru-RU). The writer stores
 * those as codes of their own (ids from 164), so the file shows what `format` says in every locale.
 */

export const BUILTIN: Readonly<Record<number, string>> = {
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '$#,##0_);($#,##0)',
  6: '$#,##0_);[Red]($#,##0)',
  7: '$#,##0.00_);($#,##0.00)',
  8: '$#,##0.00_);[Red]($#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  23: 'General',
  24: 'General',
  25: 'General',
  26: 'General',
  27: 'm/d/yy',
  28: 'm/d/yy',
  29: 'm/d/yy',
  30: 'm/d/yy',
  31: 'm/d/yy',
  32: 'h:mm:ss',
  33: 'h:mm:ss',
  34: 'h:mm:ss',
  35: 'h:mm:ss',
  36: 'm/d/yy',
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
  50: 'm/d/yy',
  51: 'm/d/yy',
  52: 'm/d/yy',
  53: 'm/d/yy',
  54: 'm/d/yy',
  55: 'm/d/yy',
  56: 'm/d/yy',
  57: 'm/d/yy',
  58: 'm/d/yy',
}

/** The built-in formats Excel shows by the regional settings. */
const span = (lo: number, hi: number) => Array.from({ length: hi - lo + 1 }, (_, i) => lo + i)

export const REGIONAL: ReadonlySet<number> = new Set([...span(5, 8), ...span(14, 36), ...span(41, 44), ...span(50, 58)])

/** Built-in ids by code, for the formats a file can store by number in every locale. */
export const BUILTIN_ID: ReadonlyMap<string, number> = new Map(
  Object.entries(BUILTIN)
    .filter(([id]) => !REGIONAL.has(Number(id)))
    .map(([id, code]) => [code, Number(id)]),
)
