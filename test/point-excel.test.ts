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
import type { PointStep } from '../bench/coverage/ui/cases.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { refAt, type PointAction } from '../src/index.js'
import './locales.js'

/** Cases where `refAt` knowingly differs from Excel, by `<set>:<id>` in every locale or `<tag>:<set>:<id>` in one, with the reason. */
const KNOWN: Record<string, string> = {
  'point-third-held:typed-6': "a sheet the book does not have (`=Sheet1!A`): Excel takes the name for another workbook's and the arrow does nothing; refAt reads text and knows no book's sheets",
  ...Object.fromEntries(
    ['en-US:point-caret-held-3:end-4', 'en-US:point-caret-held-3:inside-4'].map((k) => [
      k,
      "too few arguments once closed (=ROUND(1.5 and =ROUND(A1*2 in en-US, where ROUND is known): Excel does nothing, as Return would bring its window; refAt reads no function's arguments and enters the formula",
    ]),
  ),
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

/**
 * A case that types a text, moves the caret in Edit mode (F2, then Home and arrows, F2 again) or not, then presses Right
 * in Point mode: the text, where the caret is, and the step of the arrow; null for any other case.
 */
function arrowAt(steps: PointStep[]): { typed: string; caret: number; arrow: number } | null {
  const [first, ...rest] = steps
  if (!first || !('type' in first)) return null
  let [caret, edit] = [first.type.length, false]
  for (const [k, s] of rest.entries()) {
    if (!('key' in s) || s.shift || s.cmd) return null
    if (s.key === 'f2') edit = !edit
    else if (edit && s.key === 'home') caret = 0
    else if (edit && (s.key === 'left' || s.key === 'right')) caret = Math.min(first.type.length, Math.max(0, caret + (s.key === 'left' ? -1 : 1)))
    else if (!edit && s.key === 'right') return { typed: first.type, caret, arrow: k + 1 }
    else return null
  }
  return null
}

describe.each(SETS)('Point mode $tag $set', ({ tag, set, stored }) => {
  const arrows = stored.cases.flatMap((c) => {
    const a = c.act.do === 'point' ? arrowAt(c.act.steps) : null
    return a && c.act.do === 'point' ? [[c.id, { ...a, at: c.act.at }] as const] : []
  })
  it.each(arrows)('%s', (id, { typed, caret, arrow, at }) => {
    const bars = stored.answers[id].point!.map((s) => barOf(s.bar))
    const [before, after] = [bars[arrow - 1], bars[arrow]]
    expect(before, 'the bar shows the text typed').toEqual({ cell: at, text: typed })
    const put = (ref: string) => after.text === `${typed.slice(0, caret)}${ref}${typed.slice(caret)}`
    const excel: PointAction = after.cell !== at ? { action: 'commit' } : put('D3') ? { action: 'insert', at: caret } : put('=D3') ? { action: 'insert', at: caret, equals: true } : after.text === typed ? { action: 'none' } : (() => {
      throw new Error(`${id}: the bar after the arrow is ${after.text}`)
    })()
    const ours = refAt(typed, caret, { locale: tag })
    const known = KNOWN[`${set}:${id}`] ?? KNOWN[`${tag}:${set}:${id}`]
    if (known) expect(ours, `${typed} is known to differ: ${known}`).not.toEqual(excel)
    else expect(ours, `${typed} with the caret at ${caret}`).toEqual(excel)
  })
})
