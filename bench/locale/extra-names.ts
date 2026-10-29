/**
 * Local names of functions that Excel for Mac does not know, so that its snapshot of a locale's formula text
 * (`test/excel/locale/<tag>.json`) keeps them in English, by locale. They are Microsoft's own: the titles of the pages of
 * Excel's function reference in the locale's language (support.microsoft.com/ru-ru/excel/functions/<name>-function, and
 * for CALL and REGISTER.ID the page "Использование функций ВЫЗВАТЬ и РЕГИСТРАТОР"), read in October 2026. Excel for
 * Windows in the locale has not confirmed them. `generate.ts` writes them into the locale file's formula text beside the
 * snapshot's names, and `test/locale-excel.test.ts` compares the file with both; the English name is read there too.
 */
export const EXTRA_NAMES: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  'ru-RU': {
    CALL: 'ВЫЗВАТЬ',
    ENCODEURL: 'КОДИР.URL',
    EUROCONVERT: 'ПЕРЕСЧЕТЕВРО',
    FILTERXML: 'ФИЛЬТР.XML',
    REGEXEXTRACT: 'РЕГИЗВЛЕЧЬ',
    REGEXREPLACE: 'РЕГЗАМЕНИТЬ',
    REGEXTEST: 'РЕГТЕСТ',
    'REGISTER.ID': 'РЕГИСТРАТОР.ИД',
    WEBSERVICE: 'ВЕБСЛУЖБА',
  },
}

/** A locale's names: the snapshot's, and the extra ones for the functions it keeps in English; an error where both name one. */
export function withExtraNames(tag: string, functions: Readonly<Record<string, string>>): Record<string, string> {
  const extra = EXTRA_NAMES[tag] ?? {}
  const both = Object.keys(extra).filter((f) => f in functions && functions[f] !== f)
  if (both.length) throw new Error(`${tag}: the snapshot names ${both.join(', ')} itself: drop them from EXTRA_NAMES`)
  return { ...functions, ...extra }
}
