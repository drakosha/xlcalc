/**
 * Built-in number formats: a file stores only their number. The codes are the ones Excel en-US shows for them;
 * the date and time formats (14 to 22) and the currency ones (5 to 8, 41 to 44) Excel shows by the regional
 * settings, so a file of another locale shows them otherwise. The writer stores those as codes of their own
 * (ids from 164), so the file shows what `format` says in every locale.
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
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
}

/** The built-in formats Excel shows by the regional settings. */
export const REGIONAL = new Set([5, 6, 7, 8, 14, 15, 16, 17, 18, 19, 20, 21, 22, 41, 42, 43, 44])

/** Built-in ids by code, for the formats a file can store by number in every locale. */
export const BUILTIN_ID: ReadonlyMap<string, number> = new Map(
  Object.entries(BUILTIN)
    .filter(([id]) => !REGIONAL.has(Number(id)))
    .map(([id, code]) => [code, Number(id)]),
)
