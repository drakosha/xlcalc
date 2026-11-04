/**
 * The number format codes `writeXlsx` and `writeKept` refuse, as Excel does not open a file with them (`numfmt/refused.ts`),
 * against the answers of Excel for Windows: `bare-held`, `refused-held` and `written-held` were predicted before Excel ran them.
 * The rule is strict where it does not know: it may refuse a code Excel opens, which quotes mend, and missed none of these.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { InputError, type BookInput } from '../src/index.js'
import { formatRefused } from '../src/text.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { zip } from '../src/xlsx/zip/write.js'
import { partsOf } from './keep-files.js'

const SETS = ['digits-date-letters', 'bare-literals', 'bare-held', 'refused-held', 'written-held']
/** Short codes Excel opens that the rule refuses: nothing that quotes do not mend. */
const STRICT = ['0 шт/0', '0:', '0 :шт']
/** Past this a code's size is an estimate with a margin: a literal this long is refused where Excel may still open it. */
const LONG = 80

interface Set_ { cases: { id: string; cell: { format: string } }[]; rejected?: string[]; answers: Record<string, unknown> }
const answers = (name: string) => {
  const set = JSON.parse(readFileSync(new URL(`./excel/win/display/${name}.json`, import.meta.url), 'utf8')) as Set_
  const refused = new Set(set.rejected ?? [])
  return set.cases.filter((c) => refused.has(c.id) || set.answers[c.id]).map((c) => ({ code: c.cell.format, refused: refused.has(c.id) }))
}

describe('formatRefused', () => {
  it.each(SETS)('%s: every code Excel refused is refused, and those it opened but for the strict ones', (name) => {
    const set = answers(name)
    expect(set.length).toBeGreaterThan(30)
    expect(set.filter((c) => c.refused && !formatRefused(c.code)).map((c) => c.code)).toEqual([])
    expect(set.filter((c) => !c.refused && formatRefused(c.code) && c.code.length < LONG && !STRICT.includes(c.code)).map((c) => c.code)).toEqual([])
  })

  it.each(['0 шт', '0 шт.', '#,##0 руб', '#,##0 ₽', '#,##0.00 ₽', '#,##0,0 "млн"', '0 кг', '0 м', '0 x', '0 "kg"', '0 "шт/ч"', '_-* #,##0.00 "₽"_-;-* #,##0.00 "₽"_-;_-* "-"?? "₽"_-;_-@_-', 'dd.mm.yyyy', '[h]:mm', 'mm:ss.00', '[Red]0;[Blue]-0', '[<1]0;[<2]0;0', '[DBNum1]0', '0;-0;0;@', '# ?/?', '0.00E+00', 'General', 'General шт', '@ шт'])('%s is written', (code) =>
    expect(formatRefused(code)).toBe(false),
  )

  it.each(['0 m', '0 kg', '0 USD', '0 шт/ч', '0 л/мин', '#,##0 руб/м', 'Итог: 0', '0 samples', 'General 0', '@ General', '0 @', '@;@', '0;0;0;0', '[<1]0;[<2]0;[<3]0;0', '[<1][<2]0', '[Red][Blue]0', '[Foo]0', '0 "шт', '0\\', '0;0;0;0;0', 'ss.0000', '0'.repeat(256), `0 "${'x'.repeat(122)}"`, `"${'x'.repeat(120)}";"${'x'.repeat(120)}"`])('%s is not', (code) =>
    expect(formatRefused(code)).toBe(true),
  )
})

describe('formatRefused as the package gives it', () => {
  it('a built-in format by its number is not refused, and what is no format is an error', () => {
    for (const f of [0, 4, 14, 49]) expect(formatRefused(f)).toBe(false)
    expect(() => formatRefused(1000 as never)).toThrow(InputError)
    expect(() => formatRefused(null as never)).toThrow(InputError)
  })

  it('agrees with the writer', async () => {
    for (const numberFormat of ['0 kg', 'Итог: 0', '[Foo]0', '0 "kg"', '0 шт', '#,##0 ₽']) {
      const written = writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 1, numberFormat } } }] } }).then(() => false, () => true)
      expect(await written, numberFormat).toBe(formatRefused(numberFormat))
    }
  })
})

describe('writing a refused code', () => {
  const book = (numberFormat: string): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1, numberFormat } } }] })
  const enc = new TextEncoder()

  it('writeXlsx refuses it, naming the code and the mending', async () => {
    const refused = writeXlsx({ book: book('0 kg') })
    await expect(refused).rejects.toThrow(InputError)
    await expect(refused).rejects.toThrow('Excel does not open a file with the number format "0 kg": text in a format goes in quotes')
    expect((await writeXlsx({ book: book('0 шт') })).data.length).toBeGreaterThan(0)
  })

  it('a conditional format or a style no cell uses is refused too, as Excel reads every code on opening', async () => {
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } }, conditionalFormats: [{ type: 'expression', ranges: ['A1'], formula: 'TRUE', style: { numberFormat: '0 kg' } }] }] }
    await expect(writeXlsx({ book: b })).rejects.toThrow('0 kg')
  })

  it('writeKept refuses one a change brings, and no file comes', async () => {
    const { data } = await writeXlsx({ book: book('0') })
    const { book: read } = await openKept(data)
    const next = { ...read, sheets: [{ ...read.sheets[0], cells: { A1: { value: 1, numberFormat: '0 шт/ч' } } }] }
    await expect(writeKept(data, { book: next })).rejects.toThrow('0 шт/ч')
  })

  it("a file with one reads, and a change elsewhere leaves it the file's own", async () => {
    const parts = await partsOf((await writeXlsx({ book: book('0 "kg"') })).data)
    parts.set('xl/styles.xml', parts.get('xl/styles.xml')!.replace('formatCode="0 &quot;kg&quot;"', 'formatCode="0 kg"'))
    const data = await zip([...parts].map(([name, text]) => ({ name, data: enc.encode(text) })))
    const read = await readXlsx(data)
    expect(read.book.sheets[0].cells.A1).toMatchObject({ value: 1, numberFormat: '0 kg' })
    const { book: kept } = await openKept(data)
    expect((await writeKept(data, { book: kept })).data).toBe(data)
    const next = { ...kept, sheets: [{ ...kept.sheets[0], cells: { ...kept.sheets[0].cells, B1: { value: 2 } } }] }
    const written = await writeKept(data, { book: next })
    expect((await partsOf(written.data)).get('xl/styles.xml')).toContain('formatCode="0 kg"')
    // the workbook as a new file is another file, which takes the code anew
    await expect(writeXlsx({ book: read.book })).rejects.toThrow('0 kg')
  })
})
