/**
 * Row heights in Excel for Windows: the Mac step's books (`rowsBook`, `bench/coverage/rows.ts`), each case's row height
 * read when the book opened, its texts set through the object model (`Value2` for `value`, `Formula2Local` for `local`)
 * or pasted (`Set-Clipboard` and `Worksheet.Paste` at the cell; the keyboard's cases are the Mac's alone) and its row
 * fitted (`EntireRow.AutoFit`) when the case fits it, then each row's `RowHeight` and `UseStandardHeight` and its cells'
 * `Formula` and `WrapText` are read, and the book is saved, its `<row>` elements read back. The answers go to `test/excel/win/rows/[<tag>/]<set>.json`.
 */
import { writeFileSync } from 'node:fs'
import { writeXlsx } from '../../src/xlsx.js'
import { inParts } from '../coverage/book.js'
import { cellOf, keptRows, partAnswers, ROWS_PART, rowsFile } from '../coverage/rows.js'
import { caseCells, caseRow, rowsBook, type RowCase, type RowSet } from '../coverage/rows/cases.js'
import { answeredLine, inCaseOrder, secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { inJobs, metaOf, savedInto, type Step, type Tag, type WinEnv } from './remote.js'

export async function stepRows(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<RowSet>('rows', set)).buildRowCases(tag).filter((c) => c.via !== 'keys' && c.via !== 'fill')
  const file = rowsFile(set, tag, true)
  const { answers } = keptRows(file, cases)
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  // a column's style reaches every row of the sheet: such a case has a book of its own, as on the Mac
  const alone = (c: RowCase) => Object.values(c.columns ?? {}).some((x) => x.look)
  const parts = [...inParts(todo.filter((c) => !alone(c)), ROWS_PART), ...todo.filter(alone).map((c) => [c])]
  const book = async (part: RowCase[], f: string) => {
    const { data, issues } = await writeXlsx({ book: rowsBook(part) }, { values: 'omit' })
    if (issues.length) throw new Error(`rows ${set}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const rows = part.map((_, k) => caseRow(k))
    const acts: Step[] = part.flatMap((c, k) => [
      ...(c.typed ?? []).map(([col, text]) => (c.via === 'paste' ? { do: 'paste', sheet: 'S', at: cellOf(k, col), text } : { do: 'set', sheet: 'S', cell: cellOf(k, col), kind: c.via === 'value' ? 's' : 'l2', value: text })),
      ...(!c.typed || c.fitAfter ? [{ do: 'rowfit', sheet: 'S', range: `A${rows[k]}` }] : []),
    ])
    const cells = part.flatMap((c, k) => caseCells(c).map(([dr, col]) => cellOf(k, col, dr)))
    return { file: f, steps: [{ do: 'heights', sheet: 'S', rows }, ...acts, { do: 'calc' }, { do: 'heights', sheet: 'S', rows }, { do: 'wraps', sheet: 'S', cells }, { do: 'saveAs' }] }
  }
  const results: [RowCase[], unknown[], Uint8Array][] = []
  await inJobs(`rows-${set}-${tag}`, parts, book, savedInto(results, (part, why) => `rows ${set} ${tag}: a part did not run: ${why}`), save)
  for (const [part, got, saved] of results) {
    // the answers of the steps: heights, set or paste (null, or the error it got) and rowfit, calc, heights, wraps, saveAs
    const before = (got[0] as [number, boolean][]).map(([h]) => h)
    let i = 1
    const refused = part.map((c) => {
      const out = (c.typed ?? []).map(() => {
        const x = got[i++]
        return x && typeof x === 'object' && 'error' in x ? String((x as { error: string }).error) : ''
      })
      if (!c.typed || c.fitAfter) i++
      return out
    })
    const after = got[i + 1] as [number, boolean][]
    const wraps = got[i + 2] as [string, boolean][]
    let at = 0
    const read = part.map((c, k): [number, boolean, [string, boolean][]] => [after[k][0], after[k][1], wraps.slice(at, (at += caseCells(c).length))])
    Object.assign(answers, await partAnswers(part, before, refused, read, saved))
  }
  save()
  console.log(answeredLine(`rows ${set} ${tag}`, todo.length, answers, cases))
}
