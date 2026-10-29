/**
 * The ui sets in Excel for Windows (`bench/coverage/ui/`), each case in its own book as the Mac step writes it. Over SSH
 * Excel has no screen: a case of `bars` reads its cells' formulas through the object model alone (`Formula2`,
 * `FormulaArray`, `HasArray`, `HasSpill`, `SpillParent`), a paste puts the text on the clipboard (`Set-Clipboard`) and
 * pastes it with the cell selected (`Worksheet.Paste`), and the cases of `fill` (the keyboard's), `ribbon` and `choose`
 * (the ribbon's Number Format box) are the Mac's alone. The ribbon's buttons (`press`) need a session on the screen
 * (`screen.ts`): `press.ps1` presses them there by UI Automation, by their names in the interface's language
 * (`pressName`), answers a window Excel brings, and saves each book, which is then read as the others are, with the
 * window's text; without a session on the screen they wait. Then
 * the cells and their columns are read and the book saved, as on the Mac. The answers go to
 * `test/excel/win/ui/[<tag>/]<set>.json`; a paste Excel or the clipboard refuses stops the run with its error.
 */
import { writeFileSync } from 'node:fs'
import { writeXlsx } from '../../src/xlsx.js'
import { caseAnswer, caseFile, uiFile, type RawCell, type StoredUi } from '../coverage/ui.js'
import { clipboardText, macOnly, pressName, readCells, readColumns, uiBook, type UiCase, type UiSet } from '../coverage/ui/cases.js'
import { answeredLine, answerSaver, keptSame, loadStored, setModule } from '../coverage/stored.js'
import { inJobs, metaOf, savedInto, type Step, type Tag, type WinEnv } from './remote.js'
import { onScreen, savedOnScreen } from './screen.js'

/** The cases of the ribbon's buttons on the screen: each case's book with its buttons pressed and saved, by id; a case whose press failed is left out, with its error said. */
async function pressOnScreen(set: string, tag: Tag, cases: UiCase[]): Promise<Map<string, { saved: Uint8Array; answer: Record<string, unknown> }>> {
  const books = new Map<string, Uint8Array>()
  for (const c of cases) books.set(c.id, await caseFile(c, `${set} ${c.id}`))
  const byId = new Map(cases.map((c) => [c.id, c.act as Extract<UiCase['act'], { do: 'press' }>]))
  return savedOnScreen('press', 'press.ps1', books, (id) => {
    const act = byId.get(id)!
    return { range: act.range, buttons: act.buttons.map((b) => pressName(b, 'Windows', tag)), enter: act.enter === true }
  }, 60 + 15 * cases.length)
}

export async function stepUi(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<UiSet>('ui', set)).buildUiCases().filter((c) => !macOnly(c.act))
  const file = uiFile(set, tag, true)
  const answers = keptSame({ cases }, loadStored<StoredUi>(file))
  const t0 = Date.now()
  const save = answerSaver(file, metaOf(e, tag), true, t0, cases, answers)
  const waiting = cases.filter((c) => !answers[c.id] && c.act.do === 'press')
  const pressed = waiting.length && onScreen() ? await pressOnScreen(set, tag, waiting) : new Map<string, { saved: Uint8Array; answer: Record<string, unknown> }>()
  if (waiting.length > pressed.size) console.log(`ui ${set} ${tag}: ${waiting.length - pressed.size} cases of the ribbon's buttons wait for a session on the screen`)
  const todo = cases.filter((c) => !answers[c.id] && (c.act.do !== 'press' || pressed.has(c.id)))
  const results: [UiCase, unknown[], Uint8Array][] = []
  const book = async (c: UiCase, f: string) => {
    const { data, issues } = pressed.has(c.id) ? { data: pressed.get(c.id)!.saved, issues: [] } : await writeXlsx({ book: uiBook(c) }, { values: 'omit' })
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
    answers[c.id] = await caseAnswer(c, String(pressed.get(c.id)?.answer.dialog ?? ''), got[at + 2] as RawCell[], got[at + 3] as [number, number, boolean][], saved)
  }
  save()
  console.log(answeredLine(`ui ${set} ${tag}`, todo.length, answers, cases))
}
