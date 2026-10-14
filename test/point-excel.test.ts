/**
 * `refAt` against Excel's Point mode on the Mac's keyboard (`test/excel/ui/[ru-RU/]point-*.json`, the step `ui` of
 * `bench/coverage`): each case that types a text over C3 and presses Right, what Excel did is read from the formula
 * bar after the arrow: the active cell moved (the text was entered), the text grew by the address of D3 (a reference went
 * in), or neither (nothing happened). The keys pressed while pointing (an arrow again, Shift, Command, Home, Page Down,
 * F2) are the host's to follow, as `docs/entry.md` describes them, and are not compared here.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredUi } from '../bench/coverage/ui.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { refAt, type PointAction } from '../src/index.js'
import './locales.js'

/** Cases where `refAt` knowingly differs from Excel, by `<set>:<id>` in every locale, with the reason. */
const KNOWN: Record<string, string> = {
  'point-third-held:typed-6': "a sheet the book does not have (`=Sheet1!A`): Excel takes the name for another workbook's and the arrow does nothing; refAt reads text and knows no book's sheets",
}

/** The active cell and the text of a formula bar's description, `Formula Bar. C,3 =D3` (`Строка формул. C,3 =D3`). */
function barOf(description: string): { cell: string; text: string } {
  const m = /^[^.]*\. ([A-Z]+),(\d+) ?(.*)$/s.exec(description)
  if (!m) throw new Error(`no formula bar in ${JSON.stringify(description)}`)
  return { cell: `${m[1]}${m[2]}`, text: m[3] }
}

const SETS = TAGS.flatMap((tag) => {
  const dir = new URL(`./excel/ui/${localeDir(tag)}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.startsWith('point-') && f.endsWith('.json'))
    .map((f) => ({ tag, set: f.slice(0, -5), stored: JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredUi }))
})

describe.each(SETS)('Point mode $tag $set', ({ tag, set, stored }) => {
  const typedThenRight = stored.cases.filter((c) => c.act.do === 'point' && c.act.steps.length === 3 && 'type' in c.act.steps[0] && 'key' in c.act.steps[1] && c.act.steps[1].key === 'right' && !c.act.steps[1].shift && !c.act.steps[1].cmd)
  it.each(typedThenRight.map((c) => [c.id, c] as const))('%s', (_, c) => {
    if (c.act.do !== 'point' || !('type' in c.act.steps[0])) throw new Error('no typed text')
    const typed = c.act.steps[0].type
    const [before, after] = stored.answers[c.id].point!.map((s) => barOf(s.bar))
    expect(before, 'the bar shows the text typed').toEqual({ cell: c.act.at, text: typed })
    const excel: PointAction = after.cell !== c.act.at ? { action: 'commit' } : after.text === `${typed}D3` ? { action: 'insert', at: typed.length } : after.text === typed ? { action: 'none' } : (() => {
      throw new Error(`${c.id}: the bar after the arrow is ${after.text}`)
    })()
    const known = KNOWN[`${set}:${c.id}`]
    if (known) expect(refAt(typed, undefined, { locale: tag }), `${typed} is known to differ: ${known}`).not.toEqual(excel)
    else expect(refAt(typed, undefined, { locale: tag }), typed).toEqual(excel)
  })
})
