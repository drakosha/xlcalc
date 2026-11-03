/**
 * The form Excel 365 saves a typed formula in: a formula of the old kind (`<f>`, or a shared formula's `t="shared"`),
 * which Excel before 365 reads as it is, or a dynamic array formula (`cm` with the XLDAPR metadata, `t="array"`).
 * Excel opens a book of values in S!A1:C5, takes each case's formula typed into its range (`formula2`, as a person
 * types it), computes and saves the book (`npx tsx bench/coverage/run.ts xlsx forms` into `test/excel/xlsx/forms/`;
 * Excel for Windows through `bench/windows/run.ts xlsx forms` into `test/excel/win/xlsx/forms/`). `test/xlsx-forms.test.ts`
 * checks that `writeXlsx` writes each formula in the form Excel saved it in.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import type { BookInput } from '../../../src/contract.js'
import { compute } from '../../../src/index.js'
import { legacyConverter } from '../../../src/legacy/convert.js'
import { rectCells, parseRange } from '../../../src/structure/area.js'
import { copyFormula } from '../../../src/structure/copy.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { oldKind } from '../../../src/xlsx/write/kind.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { attributes } from '../../../src/xlsx/xml/scan.js'
import { readText } from '../../../src/xlsx/xml/text.js'

/** A formula typed into a range of sheet S (one cell, or a range Excel fills relative to its first cell). */
export interface FormCase {
  id: string
  range: string
  formula: string
}

/** Formulas of each kind: scalars, ranges where one value goes (with `@`), array functions without a spill, new functions, spills, shared formulas. */
export const FORM_CASES: FormCase[] = [
  { id: 'scalar', range: 'E1', formula: '=A1*B1' },
  { id: 'sum', range: 'E2', formula: '=SUM(A1:A5)' },
  { id: 'at-range', range: 'E3', formula: '=@A1:A5*2' },
  { id: 'at-cell', range: 'E4', formula: '=@A4+1' },
  { id: 'vlookup', range: 'E5', formula: '=VLOOKUP(3,A1:B5,2,FALSE)' },
  { id: 'sum-if', range: 'E6', formula: '=SUM(IF(A1:A5>2,B1:B5,0))' },
  { id: 'sumproduct', range: 'E7', formula: '=SUMPRODUCT(A1:A5,B1:B5)' },
  { id: 'sumproduct-test', range: 'E8', formula: '=SUMPRODUCT((A1:A5>2)*B1:B5)' },
  { id: 'max-product', range: 'E9', formula: '=MAX(A1:A5*B1:B5)' },
  { id: 'sum-index-zero', range: 'E10', formula: '=SUM(INDEX(A1:B5,0,1))' },
  { id: 'max-index-zero', range: 'E11', formula: '=MAX(INDEX(A1:A5*B1:B5,0))' },
  { id: 'sum-row', range: 'E12', formula: '=SUM(ROW(A1:A3))' },
  { id: 'countif', range: 'E13', formula: '=COUNTIF(A1:A5,">2")' },
  { id: 'sumifs', range: 'E14', formula: '=SUMIFS(B1:B5,A1:A5,">2")' },
  { id: 'index-match', range: 'E15', formula: '=INDEX(B1:B5,MATCH(3,A1:A5,0))' },
  { id: 'let', range: 'E16', formula: '=LET(x,A1,x*2)' },
  { id: 'lambda', range: 'E17', formula: '=LAMBDA(x,x*2)(A2)' },
  { id: 'xlookup', range: 'E18', formula: '=XLOOKUP(3,A1:A5,B1:B5)' },
  { id: 'textjoin', range: 'E19', formula: '=TEXTJOIN(",",TRUE,C1:C3)' },
  { id: 'sum-filter', range: 'E20', formula: '=SUM(FILTER(A1:A5,A1:A5>2))' },
  { id: 'concat', range: 'E21', formula: '=CONCAT(C1:C3)' },
  { id: 'at-index', range: 'E22', formula: '=@INDEX(A1:B5,0,1)' },
  { id: 'if', range: 'E23', formula: '=IF(A1>0,"y","n")' },
  { id: 'text', range: 'E24', formula: '=A1&"x"' },
  { id: 'let-range', range: 'E25', formula: '=LET(r,A1:A5,SUM(r))' },
  { id: 'and-range', range: 'E26', formula: '=AND(A1:A5>0)' },
  { id: 'rows', range: 'E27', formula: '=ROWS(A1:A5)' },
  { id: 'sum-times', range: 'E28', formula: '=SUM(A1:A5*1)' },
  { id: 'xlookup-row', range: 'E29', formula: '=XLOOKUP(3,A1:A5,A1:B5)' },
  { id: 'spill-sequence', range: 'G1', formula: '=SEQUENCE(3)' },
  { id: 'spill-range', range: 'H1', formula: '=A1:A3*2' },
  { id: 'spill-filter', range: 'I1', formula: '=FILTER(A1:A5,A1:A5>2)' },
  { id: 'spill-index', range: 'J1', formula: '=INDEX(A1:B5,0,1)' },
  { id: 'spill-row', range: 'K1', formula: '=ROW(A1:A3)' },
  { id: 'shared', range: 'L1:L3', formula: '=A1*B1' },
  { id: 'shared-sum-if', range: 'M1:M3', formula: '=SUM(IF(A1:A5>A1,1,0))' },
]

/** Every cell of the cases' ranges. */
export const FORM_CELLS = FORM_CASES.flatMap((c) => [...rectCells(parseRange(c.range)!)])

/** The book Excel opens: values in S!A1:C5 alone, the formulas typed in by Excel. */
export const FORMS_BOOK: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: Object.fromEntries([
        ...[1, 2, 3, 4, 5].flatMap((r) => [[`A${r}`, { value: r }], [`B${r}`, { value: r * 10 }]]),
        ...['a', 'b', 'c'].map((v, i) => [`C${i + 1}`, { value: v }]),
      ]),
    },
  ],
}

/** A formula's form in a file: of the old kind (`<f>`, a shared formula too), a dynamic array (`cm`) or an array formula of the old kind. */
export type FileForm = 'old' | 'dynamic' | 'array'

/** A cell's formula as a saved file holds it: its form, the text (none for a shared formula's copy) and a dynamic array's ref. */
export interface FormSaved {
  form: FileForm
  text?: string
  ref?: string
}

/** Each case cell's formula in the first sheet of a file, by address: `dynamic` tells whether a `cm` marks a dynamic array (one metadata in the files measured). */
export async function savedForms(data: Uint8Array): Promise<Record<string, FormSaved>> {
  const xml = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
  const out: Record<string, FormSaved> = {}
  for (const c of FORM_CASES)
    for (const addr of rectCells(parseRange(c.range)!)) {
      const m = new RegExp(`<c r="${addr}"([^>]*?)(?:/>|>(<f[^>]*?(?:/>|>[^<]*</f>))?)`).exec(xml)
      const f = m?.[2] && /^<f([^>]*?)(?:\/>|>([^<]*)<\/f>)$/.exec(m[2])
      if (!f) continue
      const fa = attributes(f[1])
      const form: FileForm = fa.t !== 'array' ? 'old' : / cm="\d+"/.test(m[1]) ? 'dynamic' : 'array'
      out[addr] = { form, ...(f[2] ? { text: readText(f[2]) } : {}), ...(form === 'dynamic' && fa.ref ? { ref: fa.ref } : {}) }
    }
  return out
}

/** The book with the case formulas in every cell of their ranges, as Excel fills a typed range (the engine's text, without `=`). */
export function formsBook(): BookInput {
  const cells = { ...FORMS_BOOK.sheets[0].cells }
  for (const c of FORM_CASES) {
    const [first, ...rest] = rectCells(parseRange(c.range)!)
    const formula = c.formula.slice(1)
    cells[first] = { formula }
    for (const a of rest) cells[a] = { formula: copyFormula(formula, { sheet: 'S', cell: first }, { sheet: 'S', cell: a }) }
  }
  return { sheets: [{ name: 'S', cells }] }
}

export const FORMS_PREDICTED = fileURLToPath(new URL('./forms.predicted.json', import.meta.url))
export const SHOWN_PREDICTED = fileURLToPath(new URL('./forms-shown.predicted.json', import.meta.url))

/**
 * The formulas Excel shows (`formula2`), predicted: in the file Excel saved, opened again, as the reader reads it
 * (`reopened`, `@` where the reader puts it into a formula of the old kind); in the file `writeXlsx` writes, the book's.
 */
export async function predictShown(saved: Uint8Array): Promise<Record<string, Record<string, string>>> {
  const read = (await readXlsx(saved)).book.sheets[0].cells
  const book = formsBook().sheets[0].cells
  const by = (cells: typeof book) => Object.fromEntries(FORM_CELLS.map((a) => [a, `=${cells[a].formula}`]))
  return { reopened: by(read), written: by(book) }
}

/**
 * The forms predicted before Excel ran, by the rule `writeKept` writes a formula by (`write/kind.ts` as it was then): a formula that
 * spills is a dynamic array over its spill; else of the old kind where reading it back as one gives the formula (`@`
 * taken out of its text), a dynamic array where not. A shared formula's copies of the old kind without their text.
 */
export async function predictForms(): Promise<Record<string, FormSaved>> {
  const book = formsBook()
  const result = compute(book)
  const written = await savedForms((await writeXlsx({ book, result })).data)
  const legacy = legacyConverter()
  const out: Record<string, FormSaved> = {}
  for (const c of FORM_CASES)
    [...rectCells(parseRange(c.range)!)].forEach((addr, i) => {
      const w = written[addr]
      const old = w.ref === addr ? oldKind(book.sheets[0].cells[addr].formula!, 'S', w.text!, legacy, (f) => f) : undefined
      if (old === undefined) out[addr] = w
      else out[addr] = i > 0 ? { form: 'old' } : { form: 'old', text: readText(old.slice(3, -4)) }
    })
  return out
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const out = await predictForms()
  if (process.argv[2] === 'shown') {
    const saved = new Uint8Array(readFileSync(fileURLToPath(new URL('../../../test/excel/xlsx/forms/forms.xlsx', import.meta.url))))
    writeFileSync(SHOWN_PREDICTED, JSON.stringify(await predictShown(saved), null, 1) + '\n')
  } else writeFileSync(FORMS_PREDICTED, JSON.stringify(out, null, 1) + '\n')
  console.log(Object.entries(out).map(([a, f]) => `${a} ${f.form} ${f.text ?? ''} ${f.ref ?? ''}`).join('\n'))
}
