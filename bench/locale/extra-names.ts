/**
 * Local names of functions that Excel for Mac does not know or keeps in English in its snapshot of a locale's formula
 * text (`test/excel/locale/<tag>.json`), by locale. Most are Microsoft's own: the titles of the pages of Excel's function
 * reference in the locale's language (support.microsoft.com/ru-ru/excel/functions/<name>-function, and for CALL and
 * REGISTER.ID the page "Использование функций ВЫЗВАТЬ и РЕГИСТРАТОР"), read in October 2026. The FORECAST functions are
 * Excel for Windows's own in ru-RU (`test/excel/win/locale/ru-RU.json`), which also confirms ENCODEURL, FILTERXML and
 * WEBSERVICE; Excel for Mac 16.113 keeps them in English. `generate.ts` writes them into the locale file's formula text
 * beside the snapshot's names, and `test/locale-excel.test.ts` compares the file with both; the English name is read
 * there too.
 */
export const EXTRA_NAMES: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  'ru-RU': {
    CALL: 'ВЫЗВАТЬ',
    ENCODEURL: 'КОДИР.URL',
    EUROCONVERT: 'ПЕРЕСЧЕТЕВРО',
    FILTERXML: 'ФИЛЬТР.XML',
    'FORECAST.ETS': 'ПРЕДСКАЗ.ETS',
    // the E of ETS is the Cyrillic Е here, as Excel for Windows writes it
    'FORECAST.ETS.CONFINT': 'ПРЕДСКАЗ.\u0415TS.ДОВИНТЕРВАЛ',
    'FORECAST.ETS.SEASONALITY': 'ПРЕДСКАЗ.ETS.СЕЗОННОСТЬ',
    'FORECAST.ETS.STAT': 'ПРЕДСКАЗ.ETS.СТАТ',
    'FORECAST.LINEAR': 'ПРЕДСКАЗ.ЛИНЕЙН',
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
