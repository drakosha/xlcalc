/**
 * The issues of `readXlsx` and `writeXlsx` in the language of `opts.locale`: a loaded locale's tag or a locale object,
 * en-US by default. Each issue says the message of the catalog in that language; its code, sheet, cells and name are the
 * same in every language, and without the option the messages are the English ones. Files with losses: those Excel
 * saved that have issues (`test/excel/xlsx/`), and one made by hand with many.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { registerLocale } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { readXlsx, writeXlsx, type XlsxIssue } from '../src/xlsx.js'
import type { XlsxMessage } from '../src/xlsx/issues.js'
import { issueText } from '../src/xlsx/messages.js'
import { pack, parts } from './xlsx-fixtures.js'

/** An issue without its message: what is the same in every language. */
const place = ({ message: _, ...rest }: XlsxIssue) => rest

/** The files Excel saved that `readXlsx` reports issues of. */
async function excelFiles(): Promise<[string, Uint8Array][]> {
  const root = new URL('./excel/xlsx/', import.meta.url)
  const out: [string, Uint8Array][] = []
  for (const dir of readdirSync(root, { withFileTypes: true }))
    if (dir.isDirectory())
      for (const f of readdirSync(new URL(`${dir.name}/`, root)).filter((x) => x.endsWith('.xlsx'))) {
        const data = readFileSync(new URL(`${dir.name}/${f}`, root))
        if ((await readXlsx(data)).issues.length) out.push([`${dir.name}/${f}`, data])
      }
  return out
}

const RICH = '<r><t>a</t></r><r><rPr><b/></rPr><t>b</t></r>'

/** A file with losses of many kinds, and each issue's code and message. */
const LOSSES = pack(
  parts({
    sheets: [
      {
        name: 'Лист',
        body:
          '<sheetPr filterMode="1"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="2000" ySplit="1000"/></sheetView></sheetViews><sheetFormatPr defaultColWidth="12" defaultRowHeight="15"/>' +
          '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>1</v></c><c r="C1" t="e"><v>#CALC!</v></c><c r="D1" t="e"><v>#N/A</v></c>' +
          '<c r="E1"><f>_xludf.MYFN(1)</f><v>0</v></c><c r="F1"><f>1+</f><v>0</v></c><c r="G1" vm="1"><v>1</v></c></row>' +
          '<row r="2"><c r="A2"><f t="dataTable" ref="A2:A2" dt2D="0" dtr="0" r1="B2"/><v>1</v></c></row><row r="3"><c r="A1048577"><v>1</v></c></row></sheetData>' +
          '<sheetProtection password="CC1A" sheet="1"/><mergeCells count="2"><mergeCell ref="A5:B5"/><mergeCell ref="B5:C5"/></mergeCells>' +
          '<conditionalFormatting sqref="A1:A3"><cfRule type="heatMap" priority="1"/></conditionalFormatting>' +
          '<dataValidations count="2"><dataValidation type="regex" sqref="B1"/><dataValidation type="whole" operator="greaterThan" imeMode="on" sqref="C1"><formula1>1</formula1></dataValidation></dataValidations>' +
          '<pageSetup orientation="landscape"/><headerFooter><oddHeader>x</oddHeader></headerFooter>',
      },
    ],
    strings: [RICH],
    styles:
      '<fonts count="2"><font><sz val="12"/><name val="Arial"/></font><font><vertAlign val="superscript"/><sz val="12"/><name val="Arial"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="darkGray"><fgColor rgb="FF00FF00"/></patternFill></fill></fills>' +
      '<borders count="1"><border/></borders><cellXfs count="2"><xf/><xf fontId="1" fillId="2"><alignment shrinkToFit="1"/></xf></cellXfs>',
    workbook: {
      before: '<workbookProtection lockStructure="1"/>',
      after:
        '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">\'Лист\'!$A$1:$B$2</definedName><definedName name="ставка" hidden="1">0.1</definedName></definedNames>' +
        '<calcPr calcMode="manual" fullPrecision="0" iterate="1" iterateCount="0" iterateDelta="-1"/>',
    },
  }),
)

const LOST: [XlsxIssue['code'], XlsxMessage][] = [
  ['defaultFont', { key: 'defaultFont', name: 'Arial', size: 12 }],
  ['calculation', { key: 'precisionAsDisplayed' }],
  ['calculation', { key: 'iterationCount', count: '0' }],
  ['calculation', { key: 'iterationChange', change: '-1' }],
  ['protection', { key: 'workbookProtection' }],
  ['name', { key: 'builtInName' }],
  ['richText', { key: 'richText' }],
  ['style', { key: 'styleLost', part: 'superscript' }],
  ['style', { key: 'styleLost', part: 'patternFill' }],
  ['style', { key: 'styleLost', part: 'shrinkToFit' }],
  ['errorValue', { key: 'errorUnspelled' }],
  ['errorConstant', { key: 'errorTyped' }],
  ['userFunction', { key: 'userFunction' }],
  ['formula', { key: 'formulaUnparsed' }],
  ['cellMetadata', { key: 'valueMetadata' }],
  ['dataTable', { key: 'dataTable' }],
  ['part', { key: 'cellOutside' }],
  ['part', { key: 'mergeOverlap' }],
  ['autoFilter', { key: 'advancedFilter' }],
  ['frozenPanes', { key: 'splitPanes' }],
  ['defaultWidth', { key: 'defaultWidth', width: 12 }],
  ['protection', { key: 'sheetPassword' }],
  ['print', { key: 'printSettings' }],
  ['conditionalFormat', { key: 'condType', type: 'heatMap' }],
  ['dataValidation', { key: 'ruleType', type: 'regex' }],
  ['dataValidation', { key: 'ruleIme' }],
]

/** Each issue's message by its code and English text: the message of `LOST` it says. */
function expected(issues: XlsxIssue[]): (XlsxMessage | undefined)[] {
  return issues.map((i) => LOST.find(([code, m]) => code === i.code && issueText('en-US', m) === i.message)?.[1])
}

const LONG = 'x'.repeat(32_768)

/** A book whose file loses something of each kind the writer reports here. */
const WRITTEN = {
  sheets: [
    {
      name: 'S',
      cells: { A1: { formula: '1+1', meta: { id: 1 } }, A2: { value: LONG, note: { text: LONG } }, A3: { value: 1, link: { url: 'https://example.com', tooltip: 'y'.repeat(256) } }, A4: { formula: 'Нет!A1' } },
      validations: [{ range: 'B1', type: 'whole' as const, operator: 'greaterThan' as const, formula1: '0', prompt: 'p'.repeat(256) }],
    },
  ],
  names: [{ name: 'rate', ref: '1+' }],
}

describe('the issues of a file read, in the language of opts.locale', () => {
  it('a locale object needs no loading; the tag of a locale not loaded, an unknown option and options that are no object are refused', async () => {
    const data = await LOSSES
    expect((await readXlsx(data, { locale: ruRU })).issues.find((i) => i.code === 'richText')?.message).toBe('Форматированный текст прочитан как простой')
    await expect(readXlsx(data, { locale: 'ru-RU' })).rejects.toThrow("opts.locale: the locale ru-RU is not loaded: import { ruRU } from 'xlcalc/locales/ru-RU' and call registerLocale(ruRU), or await loadLocale('ru-RU') from 'xlcalc/locales' (loaded: en-US)")
    await expect(readXlsx(data, { lang: 'ru-RU' } as never)).rejects.toThrow('opts: unknown field "lang" (expected locale, limits)')
    await expect(readXlsx(data, 'ru-RU' as never)).rejects.toThrow("opts: expected options such as { locale: 'ru-RU' }, got the string \"ru-RU\"")
    await expect(writeXlsx({ book: WRITTEN }, { locale: 'ru-RU' })).rejects.toThrow('opts.locale: the locale ru-RU is not loaded')
    registerLocale(ruRU)
  })

  it('every issue of a file with many losses in ru-RU: the catalog\'s message in Russian, the rest as in en-US; without the option the English one', async () => {
    const data = await LOSSES
    const [plain, en, ru, object] = await Promise.all([readXlsx(data), readXlsx(data, { locale: 'en-US' }), readXlsx(data, { locale: 'ru-RU' }), readXlsx(data, { locale: ruRU })])
    expect(en).toEqual(plain)
    expect(object).toEqual(ru)
    expect(ru.book).toEqual(plain.book)
    expect(ru.issues.map(place)).toEqual(plain.issues.map(place))
    const said = expected(plain.issues)
    expect(plain.issues.filter((_, k) => !said[k]).map((i) => `${i.code}: ${i.message}`)).toEqual([])
    expect(said.map((m) => m!.key).sort()).toEqual(LOST.map(([, m]) => m.key).sort())
    expect(ru.issues.map((i) => i.message)).toEqual(said.map((m) => issueText('ru-RU', m!)))
    expect(ru.issues.find((i) => i.code === 'calculation' && i.message.startsWith('Предельное'))).toEqual({ code: 'calculation', message: 'Предельное число итераций «0» вне пределов от 1 до 32 767: прочитано значение по умолчанию, 100' })
    expect(ru.issues.filter((i) => i.code === 'style').map((i) => [i.message, i.cells])).toEqual([
      ['Надстрочный шрифт в стиле не поддерживается', ['B1']],
      ['Заливка узором в стиле не поддерживается', ['B1']],
      ['Автоподбор ширины в стиле не поддерживается', ['B1']],
    ])
    // the hidden name is read hidden, without an issue; the built-in one is left out
    expect(ru.book.names).toEqual([{ name: 'ставка', ref: '0.1', hidden: true }])
    expect(ru.issues.filter((i) => i.code === 'name').map((i) => [i.name, i.message])).toEqual([['_xlnm.Print_Area', 'Встроенное имя (область печати, печатаемые заголовки) не поддерживается']])
  })

  it('the files Excel saved that have issues: in ru-RU each message is Russian, with the same code, sheet, cells and name', async () => {
    const files = await excelFiles()
    expect(files.length).toBeGreaterThanOrEqual(15)
    for (const [name, data] of files) {
      const [en, ru] = await Promise.all([readXlsx(data), readXlsx(data, { locale: 'ru-RU' })])
      expect(ru.issues.map(place), name).toEqual(en.issues.map(place))
      for (const [k, i] of ru.issues.entries()) {
        expect(i.message, name).toMatch(/^\p{Lu}.*\p{Script=Cyrillic}/u)
        expect(i.message, name).not.toBe(en.issues[k].message)
      }
    }
  }, 60_000)
})

describe('the issues of a file written, in the language of opts.locale', () => {
  it('each in Russian in ru-RU, the rest as in en-US; without the option the English one', async () => {
    const [plain, en, ru] = await Promise.all([writeXlsx({ book: WRITTEN }), writeXlsx({ book: WRITTEN }, { locale: 'en-US' }), writeXlsx({ book: WRITTEN }, { locale: 'ru-RU', values: 'omit' })])
    expect(en).toEqual(plain)
    expect(ru.data).toEqual(plain.data)
    const said: XlsxMessage[] = [
      { key: 'cellMeta' },
      { key: 'textCut' },
      { key: 'formulaLeftOut' },
      { key: 'tooltipCut' },
      { key: 'ruleTextCut', field: 'prompt', max: 255 },
      { key: 'noteCut' },
      { key: 'noResult' },
      { key: 'nameBadRef' },
    ]
    expect(plain.issues.map((i) => i.message)).toEqual(said.map((m) => issueText('en-US', m)))
    // `values: 'omit'` asked for: no values of formulas, and nothing reported of it
    expect(ru.issues.map(place)).toEqual(plain.issues.filter((i) => i.code !== 'noCachedValues').map(place))
    expect(ru.issues.map((i) => i.message)).toEqual(said.filter((m) => m.key !== 'noResult').map((m) => issueText('ru-RU', m)))
    expect(ru.issues.find((i) => i.cells?.[0] === 'B1')?.message).toBe('Сообщение для ввода правила проверки данных длиннее 255 знаков, предела Excel, обрезано')
    expect((await writeXlsx({ book: WRITTEN }, { locale: ruRU })).issues.find((i) => i.code === 'noCachedValues')?.message).toBe('Результат не передан: файл не хранит значения формул, Excel вычислит их при открытии файла')
  })

  it('a locale object without messages of its own speaks as the loaded locale of its tag, else in English; one without those of xlsx gives the English ones', async () => {
    const book = { sheets: [{ name: 'S', cells: { A1: { formula: '1+1' } } }] }
    const message = async (locale: typeof ruRU) => (await writeXlsx({ book }, { locale })).issues[0].message
    expect(await message({ ...ruRU, messages: undefined })).toBe('Результат не передан: файл не хранит значения формул, Excel вычислит их при открытии файла')
    expect(await message({ ...ruRU, tag: 'xx-YY', messages: undefined })).toBe('no result was passed: the file stores no formula values, and Excel calculates them when it opens the file')
    const engineOnly = { ...ruRU, tag: 'xx-ZZ', messages: { lastSheet: () => 'Последний лист' } }
    registerLocale(engineOnly)
    expect(await message(engineOnly)).toBe('no result was passed: the file stores no formula values, and Excel calculates them when it opens the file')
    expect((await writeXlsx({ book }, { locale: 'xx-ZZ' })).issues[0].message).toBe('no result was passed: the file stores no formula values, and Excel calculates them when it opens the file')
  })
})
