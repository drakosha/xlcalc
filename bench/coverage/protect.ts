/**
 * Protection against Excel for Mac en-US: `npx tsx bench/coverage/run.ts protect <set>` sends the file of each case of
 * `protect/<set>.ts` (`caseFile`), Excel opens it, takes the case's steps, reads every sheet's protection and the listed
 * cells' locked and formula hidden, and saves the book (`excel-protect.applescript`). The answers go to
 * `test/excel/protect/<set>.json` and the saved files to `test/excel/protect/<set>/<id>.xlsx`;
 * `test/protection-excel.test.ts` compares them with the engine without Excel. `FPROTECT_ONLY=id,id` runs only those
 * cases. `npx tsx bench/coverage/protect/predict.ts <set>` writes the engine's answers before Excel runs a set.
 */
import { join } from 'node:path'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { ALLOWS, caseFile, stepArgs, type Allow, type CellGuard, type ProtectAnswer, type ProtectCase, type ProtectSet, type SheetGuard } from './protect/cases.js'
import { fileAnswers, plain, runFileCases, secondsSince, setModule } from './stored.js'
import { caseFiles } from './xlsx/sent.js'
import { HERE, ROOT } from './paths.js'

const SCRIPT = join(HERE, 'excel-protect.applescript')

type RawSheet = [name: string, contents: boolean, objects: boolean, scenarios: boolean, selection: string, allow: boolean[]]
type RawCell = [address: string, locked: boolean, hidden: boolean]

/** One case in Excel: its answer and the saved file, or the repair window's text. */
async function runCase(c: ProtectCase, data: Uint8Array) {
  const r = await runSavingScript(c.id, SCRIPT, data, [JSON.stringify((c.steps ?? []).map(stepArgs)), JSON.stringify(c.read)])
  if ('rejected' in r) return r
  const raw = JSON.parse(r.out) as { error: string; sheets: RawSheet[]; cells: [string, RawCell[]][] }
  const sheets: Record<string, SheetGuard> = {}
  for (const [name, contents, objects, scenarios, selection, allow] of raw.sheets)
    sheets[name] = { contents, objects, scenarios, selection, allow: Object.fromEntries(ALLOWS.map((k, i) => [k, allow[i]])) as Record<Allow, boolean> }
  const cells: ProtectAnswer['cells'] = {}
  for (const [sheet, got] of raw.cells) cells[sheet] = { ...cells[sheet], ...Object.fromEntries(got.map(([a, locked, hidden]): [string, CellGuard] => [plain(a), { locked, hidden }])) }
  const answer: ProtectAnswer = { ...(raw.error ? { error: raw.error } : {}), sheets, cells }
  return { answer, saved: r.saved }
}

/** Runs the set's cases that have no stored answer for the same case and file and stores the answers. */
export async function stepProtect(set: string): Promise<void> {
  const built = (await setModule<ProtectSet>('protect', set)).buildProtectCases()
  const { files, sent } = await caseFiles(built.cases, (c) => caseFile(built, c))
  const store = fileAnswers<ProtectCase, ProtectAnswer>(join(ROOT, `test/excel/protect/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ version: excelVersion(), locale: 'en-US', seconds: secondsSince(t0) }, built)
  await ensureExcel('en-US')
  await runFileCases(`protect ${set}`, store, store.todo, 'FPROTECT_ONLY', (c) => runCase(c, files.get(c.id)!), save, t0)
}
