/**
 * The conditional formats the model does not read, carried by `writeKept`, against Excel (`bench/coverage/keep/rules.ts`):
 * Excel for Windows opened and saved the books (`test/excel/win/keep/rules-made/`) and did each case itself over the
 * saved `mixed` (`keep/rules*`), and Excel for Mac opened each file `writeKept` wrote for the cases (`keep/rules-open`).
 * The rules `writeKept` writes for a case are in Excel's order, with Excel's ranges and formulas, but where `KNOWN` says
 * why not; `writeKept` counts priorities from 1, where Excel keeps the numbers rules had.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { RULES_BOOKS, RULES_SETS, rulesFile, rulesOf, rulesWritten } from '../bench/coverage/keep/rules.js'

const WIN = 'test/excel/win/keep'
const file = (path: string) => new Uint8Array(readFileSync(path))
const mixed = file(`${WIN}/rules-made/mixed.xlsx`)

/** Where Excel's rules after a case are not those `writeKept` writes, and why. */
const KNOWN: Record<string, string> = {
  'last-1': "Excel's Priority puts the rule before the carried rule that had that priority; the workbook's list cannot say which side of a carried rule a rule goes, and the carried rule stays after the rule it came after",
  'held-priority-4': 'as last-1',
  'held-move-part-b': "part of a carried rule's cells moved: no carrier sees it, the rule keeps its cells, where Excel takes the moved ones along",
}

/** The cases whose rules `writeKept` writes in another order since Excel for Mac opened its files (the order of rules after deleted lines, `cond/order`): their rules compare without their order. */
const REORDERED = new Set(['held-delete-rows-top', 'held2-delete-cols-c'])

/** The cases `writeKept` refuses, by the code of the refusal: a carried rule moved to another sheet, which Excel moves there. */
const REFUSED: Record<string, string> = { 'held-move-h-other': 'moved' }

describe('the books Excel for Windows saved', () => {
  it("keep every rule in its order, a data bar's threshold without its value saved as 0 (which the model then reads)", async () => {
    for (const id of Object.keys(RULES_BOOKS)) {
      const sent = (await rulesOf(await rulesFile(id))).rules.map((r) => r.replace('dataBar E2:E21 unread', 'dataBar E2:E21 =0'))
      expect((await rulesOf(file(`${WIN}/rules-made/${id}.xlsx`))).rules, id).toEqual(sent)
    }
  })
})

describe.each(Object.entries(RULES_SETS))('the set %s', (set, cases) => {
  const answers = JSON.parse(readFileSync(`${WIN}/${set}.json`, 'utf8')) as { answers: Record<string, { error?: string }> }

  it.each(cases.map((c) => [c.id, c] as const))('%s: the rules Excel saved, as writeKept writes them', async (id, c) => {
    expect(answers.answers[id].error).toBeUndefined()
    const excel = (await rulesOf(file(`${WIN}/${set}/${id}.xlsx`))).rules
    const w = await rulesWritten(mixed, c)
    if (REFUSED[id]) return void expect('refused' in w && w.refused).toContain('another sheet')
    if (!('data' in w)) throw new Error(w.refused)
    const ours = await rulesOf(w.data)
    expect(ours.priorities).toEqual(ours.rules.map((_, i) => i + 1))
    if (KNOWN[id]) expect(ours.rules).not.toEqual(excel)
    else expect(ours.rules).toEqual(excel)
  })
})

describe('the files writeKept wrote, opened by Excel for Mac', () => {
  const opened = JSON.parse(readFileSync('test/excel/keep/rules-open.json', 'utf8')) as { cases: Record<string, { rejected?: string; refs?: { rules: string[] } }> }

  it('open without the repair window and keep their rules, a 2010 colour scale left without a reference to another sheet saved as 2007', async () => {
    const all = Object.values(RULES_SETS).flat().filter((c) => !REFUSED[c.id])
    expect(Object.keys(opened.cases).sort()).toEqual(all.map((c) => c.id).sort())
    for (const c of all) {
      const w = await rulesWritten(mixed, c)
      if (!('data' in w)) throw new Error(w.refused)
      const ours = (await rulesOf(w.data)).rules.map((r) => (c.id === 'delete-sheet-p' ? r.replace('x14 colorScale', 'colorScale') : r))
      expect(opened.cases[c.id].rejected, c.id).toBeUndefined()
      const saved = opened.cases[c.id].refs?.rules
      if (!REORDERED.has(c.id)) expect(saved, c.id).toEqual(ours)
      else {
        expect(saved, c.id).not.toEqual(ours)
        expect([...(saved ?? [])].sort(), c.id).toEqual([...ours].sort())
      }
    }
  })
})
