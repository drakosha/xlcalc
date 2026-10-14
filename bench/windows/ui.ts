/**
 * The ui sets in Excel for Windows (`bench/coverage/ui/`), each case in its own book as the Mac step writes it. Over SSH
 * Excel has no screen: a case of `bars` reads its cells' formulas through the object model alone (`Formula2`,
 * `FormulaArray`, `HasArray`, `HasSpill`, `SpillParent`), a paste puts the text on the clipboard (`Set-Clipboard`) and
 * pastes it with the cell selected (`Worksheet.Paste`), and the cases of `fill` (the keyboard's), `ribbon` and `choose`
 * (the ribbon's Number Format box, which Excel shows on a screen alone) are the Mac's alone. Then
 * the cells and their columns are read and the book saved, as on the Mac. The answers go to
 * `test/excel/win/ui/[<tag>/]<set>.json`; a paste Excel or the clipboard refuses stops the run with its error.
 */
import { writeFileSync } from 'node:fs'
import { writeXlsx } from '../../src/xlsx.js'
import { caseAnswer, uiFile, type RawCell, type StoredUi } from '../coverage/ui.js'
import { clipboardText, macOnly, readCells, readColumns, uiBook, type UiCase, type UiSet } from '../coverage/ui/cases.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { inJobs, metaOf, savedInto, type Step, type Tag, type WinEnv } from './remote.js'

export async function stepUi(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<UiSet>('ui', set)).buildUiCases().filter((c) => !macOnly(c.act))
  const file = uiFile(set, tag, true)
  const answers = keptSame({ cases }, loadStored<StoredUi>(file))
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  const results: [UiCase, unknown[], Uint8Array][] = []
  const book = async (c: UiCase, f: string) => {
    const { data, issues } = await writeXlsx({ book: uiBook(c) }, { values: 'omit' })
    if (issues.length) throw new Error(`ui ${set} ${c.id}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const paste: Step[] = c.act.do === 'paste' ? [{ do: 'paste', sheet: 'S', at: c.act.at, text: clipboardText(c.act) }] : []
    return { file: f, steps: [{ do: 'full' }, ...paste, { do: 'calc' }, { do: 'cells', sheet: 'S', cells: readCells(c) }, { do: 'widths', sheet: 'S', columns: readColumns(c) }, { do: 'saveAs' }] }
  }
  await inJobs(`ui-${set}-${tag}`, todo, book, savedInto(results, (c, why) => `ui ${set} ${tag} ${c.id}: the book did not run: ${why}`), save)
  for (const [c, got, saved] of results) {
    // the steps' answers: full, the paste (null, or its error), calc, cells, widths, saveAs
    const at = c.act.do === 'paste' ? 1 : 0
    const pasted = got[at] as { error?: string } | null
    if (at && pasted?.error) throw new Error(`ui ${set} ${tag} ${c.id}: the paste failed: ${pasted.error}`)
    answers[c.id] = await caseAnswer(c, '', got[at + 2] as RawCell[], got[at + 3] as [number, number, boolean][], saved)
  }
  save()
  console.log(answeredLine(`ui ${set} ${tag}`, todo.length, answers, cases))
}
