/**
 * The conditional formats of a sheet the reader did not all take (`readXlsx`'s `conditionalFormat` issues). A rule the
 * model does not have is carried: its carriers (`ruleItems`) move its ranges as Excel moves a rule's (lines inserted
 * right after them join them, a range ending on the sheet's last line keeps ending there, a rule with none left goes)
 * and its formulas as a cell's at its first cell. When the sheet's rules are written again (`keptRules`), each such rule
 * goes in as the file has it, with its moved ranges and formulas, after the rule of the workbook it came after in the
 * file, its priority counted among theirs. A rule the model read without a part of its look keeps the file's differential
 * format while the workbook's look of it is the same. What cannot go as Excel would is refused (`rules`).
 */
import type { ConditionalFormat } from '../../contract.js'
import { GRADED, lateAfterDeletion, lateBy, lostRefs } from '../../cond/structure.js'
import { MAX_COLS, MAX_ROWS } from '../../model.js'
import { cellCount, cornerOf, type Rect } from '../../structure/area.js'
import type { Ast } from '../../syntax/ast.js'
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { canonical, mapAst, quoteSheet, sheetKey } from '../../syntax/print.js'
import { keyOf } from '../../style-table.js'
import { toFile } from '../formulas/to-file.js'
import { Issues } from '../issues.js'
import { readConditional, type RuleSource } from '../read/conditional.js'
import { parseFormula } from '../read/formulas.js'
import { rulesText } from '../read/sheet.js'
import type { Palette } from '../styles/color.js'
import type { Dxf } from '../styles/dxf.js'
import { block, block14, x14Ext, type ConditionalWritten } from '../write/conditional.js'
import { attributes, placed } from '../xml/scan.js'
import { escapeAttr, escapeText, readText } from '../xml/text.js'
import type { Carrier, Edit, Item } from './carriers.js'
import { absText, areasOf, areaText, dollarsOf, lastCell, plainArea, sqrefCarrier, sqrefRects, type Dollars } from './refs.js'
import { edited, KeptRefusal } from './refusals.js'
import { setAttrs, tagEnd, valueAt } from './xml.js'

/** What a rule the reader read may lose and still be written as the workbook has it: a part of its look the file keeps, the deviations Excel drops too. */
const KEPT_LOSSES = new Set(['condStyleLost', 'condAverage'])

/** A rule the model does not have, to be carried. */
const unread = (s: RuleSource) => s.index === undefined && !s.pivot

/** The file's rules of a sheet (its part's text) as the reader meets them, where they lie in the part. */
export function fileRules(xml: string, sheet: string, dxfs: Dxf[], palette: Palette): RuleSource[] {
  if (!xml.includes('cfRule')) return []
  const close = /<\/(?:[\w.-]+:)?sheetData\s*>/.exec(xml)
  const at = close ? close.index + close[0].length : 0
  const { own, extensions, extAt } = rulesText(xml.slice(at))
  const out: RuleSource[] = []
  readConditional(own, extensions, sheet, dxfs, palette, new Issues(), {}, out)
  return out.map((s) => {
    const d = at + (s.x14 ? extAt : 0)
    return { ...s, start: s.start + d, end: s.end + d, block: { ...s.block, start: s.block.start + d, end: s.block.end + d } }
  })
}

/** A formula of a rule as its carrier holds it: each reference's `$`s and whether it named no sheet, in the order the tree gives them (null for a lost one). */
interface Plan {
  refs: ({ dollars: Dollars; own: boolean } | null)[]
  relative: boolean
}

/** The references of a tree in order, a lost one (`#REF!`) too, through `f`. */
const refsIn = (ast: Ast, f: (n: Ast, i: number) => Ast | undefined): Ast => {
  let i = 0
  return mapAst(ast, (n) => (n.k === 'ref' || n.k === 'ref3d' || (n.k === 'err' && n.code === 'REF') ? f(n, i++) : undefined))
}

/** A rule's formula as a carrier: every reference absolute at the rule's first cell and on its sheet; undefined without one, null when it does not parse. */
function formulaCarrier(text: string, sheet: string): { ref: string; plan: Plan } | undefined | null {
  const f = parseFormula(text)
  if (!f.ast) return null
  const plan: Plan = { refs: [], relative: false }
  const ast = refsIn(f.ast, (n) => {
    if (n.k === 'err') return void plan.refs.push(null)
    if (n.k !== 'ref' && n.k !== 'ref3d') return undefined
    const dollars = dollarsOf(n.ref)
    plan.relative ||= !dollars.c1 || !dollars.r1 || !dollars.c2 || !dollars.r2
    plan.refs.push({ dollars, own: n.k === 'ref' && n.sheet === null })
    const ref = { ...n.ref, text: absText(n.ref) }
    return n.k === 'ref' ? { ...n, sheet: n.sheet ?? sheet, ref } : { ...n, ref }
  })
  return plan.refs.some(Boolean) ? { ref: canonical(ast), plan } : undefined
}

/** A carrier's formula written back as a file's, its references with the `$`s they had and its own sheet's without the sheet; null when it lost the shape of its references. */
function formulaBack(now: string, plan: Plan, sheet: string): string | null {
  let ast: Ast
  try {
    ast = parse(now)
  } catch (e) {
    if (e instanceof FormulaError) return null
    throw e
  }
  let n = 0
  ast = refsIn(ast, (x, i) => {
    n++
    const p = plan.refs[i]
    // a reference of its own sheet lost is `#REF!`, as Excel writes it there
    if (x.k === 'err') return p?.own ? { k: 'err', code: x.code } : undefined
    if (!p || (x.k !== 'ref' && x.k !== 'ref3d')) return undefined
    const ref = { ...x.ref, text: areaText(x.ref, p.dollars) }
    return x.k === 'ref' ? { ...x, sheet: p.own && x.sheet !== null && sheetKey(x.sheet) === sheetKey(sheet) ? null : x.sheet, ref } : { ...x, ref }
  })
  return n === plan.refs.length ? toFile(ast) : null
}

/** The formulas of a `cfRule` at `start`–`end` of a part: the content of `<formula>` and `<xm:f>`, and a threshold's `val`. */
function formulasOf(xml: string, start: number, end: number): { at: [number, number]; text: string; attr: boolean }[] {
  const text = xml.slice(start, end)
  const out: { at: [number, number]; text: string; attr: boolean }[] = []
  for (const local of ['formula', 'f'])
    for (const p of placed(text, local)) {
      if (p.body === null) continue
      const open = start + text.lastIndexOf('</', p.end - 1) - p.body.length
      out.push({ at: [open, open + p.body.length], text: readText(p.body), attr: false })
    }
  for (const p of placed(text, 'cfvo')) {
    const at = valueAt(xml, { start: start + p.start }, 'val')
    if (at) out.push({ at, text: readText(xml.slice(...at)), attr: true })
  }
  return out.sort((a, b) => a.at[0] - b.at[0])
}

/** Where a block's ranges lie in the part: its `sqref`, or Excel 2010's `<xm:sqref>`. */
function sqrefAt(xml: string, s: RuleSource): [number, number] | undefined {
  if (!s.x14) return valueAt(xml, { start: s.block.start }, 'sqref')
  const text = xml.slice(s.block.start, s.block.end)
  const q = placed(text, 'sqref').at(-1)
  if (!q || q.body === null) return undefined
  const open = s.block.start + text.lastIndexOf('</', q.end - 1) - q.body.length
  return [open, open + q.body.length]
}

/** The items of the sheet's rules the model does not have, a block of the file each, and the reasons a rule's references cannot be read. */
export function ruleItems(xml: string, sources: RuleSource[], sheet: string, on: number, part: string, carrier: (ref: string, kind: Carrier['kind']) => Carrier): { items: Item[]; unparsed: string[] } {
  const items: Item[] = []
  const unparsed: string[] = []
  const blocks = new Map<number, RuleSource[]>()
  for (const s of sources) if (unread(s)) (blocks.get(s.block.start) ?? blocks.set(s.block.start, []).get(s.block.start)!).push(s)
  for (const rules of blocks.values()) {
    const first = rules[0]
    const span = sqrefAt(xml, first)
    const sqref = span ? readText(xml.slice(...span)) : first.block.sqref
    const rects = span ? sqrefRects(sqref) : null
    const what = `an unread conditional format of '${sheet}' (${sqref})`
    const formulas = rules.flatMap((s) => formulasOf(xml, s.start, s.end).map((f) => ({ ...f, carried: formulaCarrier(f.text, sheet) })))
    if (!rects || formulas.some((f) => f.carried === null)) {
      unparsed.push(what)
      continue
    }
    const plans = formulas.flatMap((f) => (f.carried ? [{ ...f, carried: f.carried }] : []))
    const relative = plans.some((f) => f.carried.plan.relative)
    const next = (axis: 'r' | 'c') => rects.flatMap((x, i) => ((axis === 'r' ? x.r2 < MAX_ROWS - 1 : x.c2 < MAX_COLS - 1) ? [{ i, x: axis === 'r' ? { ...x, r1: x.r2 + 1, r2: x.r2 + 1 } : { ...x, c1: x.c2 + 1, c2: x.c2 + 1 } }] : []))
    const [below, beside] = [next('r'), next('c')]
    const origin = cornerOf(rects)
    const union = (xs: Rect[]) => sqrefCarrier(sheet, xs) ?? lastCell(sheet)
    const carriers = [
      carrier(union(rects), 'place'),
      carrier(union(below.map((b) => b.x)), 'place'),
      carrier(union(beside.map((b) => b.x)), 'place'),
      carrier(lastCell(sheet), 'ref'),
      carrier(`${quoteSheet(sheet)}!${absText({ r1: origin.r, r2: origin.r, c1: origin.c, c2: origin.c })}`, 'place'),
      ...plans.map((f) => carrier(f.carried.ref, 'ref')),
    ]
    items.push({
      what,
      on,
      carriers,
      fix([areasNow, belowNow, besideNow, sheetNow, originNow, ...formulasNow]) {
        const own = areasOf(sheetNow)?.[0]?.sheet
        const areas = areasOf(areasNow)
        if (!areas || own === undefined) return { refuse: `${what} is no reference after the change`, code: 'rules' }
        if (areas.some((a) => 'rect' in a && sheetKey(a.sheet) !== sheetKey(own))) return { refuse: `${what} moved to another sheet, which writeKept does not write`, code: 'moved' }
        const now = areas.map((a) => ('rect' in a ? { ...a.rect } : null))
        // lines inserted right after a range join it, as Excel joins them; a range ending on the sheet's last line keeps ending there
        const join = (text: string, list: { i: number }[], axis: 'r' | 'c') =>
          areasOf(text)?.forEach((a, k) => {
            const x = now[list[k]?.i ?? -1]
            if (!x || !('rect' in a)) return
            if (axis === 'r' && a.rect.r1 > x.r2 + 1) x.r2 = a.rect.r1 - 1
            if (axis === 'c' && a.rect.c1 > x.c2 + 1) x.c2 = a.rect.c1 - 1
          })
        if (below.length) join(belowNow, below, 'r')
        if (beside.length) join(besideNow, beside, 'c')
        rects.forEach((x, i) => {
          const y = now[i]
          if (y && x.r2 === MAX_ROWS - 1) y.r2 = MAX_ROWS - 1
          if (y && x.c2 === MAX_COLS - 1) y.c2 = MAX_COLS - 1
        })
        const left = now.filter((x): x is Rect => x !== null)
        if (!left.length) return { edits: rules.map((s) => ({ part, start: s.start, end: s.end, text: '' })) }
        const edits: Edit[] = []
        const text = left.map(plainArea).join(' ')
        if (text !== sqref) edits.push({ part, start: span![0], end: span![1], text: first.x14 ? escapeText(text) : escapeAttr(text) })
        // a formula with references relative to the first cell is written for the cell it was, which is first still
        const o = areasOf(originNow)?.[0]
        const corner = cornerOf(left)
        if (relative && (!o || !('rect' in o) || o.rect.r1 !== corner.r || o.rect.c1 !== corner.c))
          return { refuse: `${what} lost its first cell, and its formulas have relative references, which writeKept does not make for the new first cell`, code: 'rules' }
        for (const [k, f] of plans.entries()) {
          if (formulasNow[k] === f.carried.ref) continue
          const back = formulaBack(formulasNow[k], f.carried.plan, own)
          if (back === null) return { refuse: `${what} has a formula whose references the change took apart`, code: 'rules' }
          edits.push({ part, start: f.at[0], end: f.at[1], text: f.attr ? escapeAttr(back) : escapeText(back) })
        }
        return { edits }
      },
    })
  }
  return { items, unparsed }
}

/** A rule as it counts for whether it is the same one: all but its ranges and formulas. */
const shape = (rule: ConditionalFormat): string => keyOf(JSON.parse(JSON.stringify(rule, (k, v: unknown) => (k === 'ranges' || k === 'formula' || k === 'formulas' || k === 'value' ? undefined : v))))

/**
 * Where each rule of `r0` went in `r1`: the longest run of the same rules in order, a rule moved after the others left
 * out (Excel's order after an operation), and between two such the rules left paired in turn (one changed); by rule of
 * `r0` its place in `r1` (-1: gone) and whether it is the same rule.
 */
function align(r0: ConditionalFormat[], r1: ConditionalFormat[]): { to: number[]; same: boolean[] } {
  const [a, b] = [r0.map(shape), r1.map(shape)]
  const to = a.map(() => -1)
  const same = a.map(() => false)
  const [n, m] = [a.length, b.length]
  if (n * m <= 4_000_000) {
    const dp = Array.from({ length: n + 1 }, () => new Int32Array(m + 1))
    for (let i = 1; i <= n; i++) for (let j = 1; j <= m; j++) dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1])
    // a tie leaves out the later rule of `r1` (a rule Excel put back after the others), or of `r0` (one changed, where
    // rules alike stay by their places): of the two the one whose rules left out of `r1` lie nearest its ends (Excel
    // moves a rule first or last), then whose rules moved the least
    const back = (later: boolean) => {
      const out = a.map(() => -1)
      for (let i = n, j = m; i > 0 && j > 0; ) {
        if (a[i - 1] === b[j - 1]) out[--i] = --j
        else if (later ? dp[i][j - 1] >= dp[i - 1][j] : dp[i][j - 1] > dp[i - 1][j]) j--
        else i--
      }
      return out
    }
    const moves = (xs: number[]) => xs.reduce((sum, j, i) => sum + (j < 0 ? 0 : Math.abs(j - i)), 0)
    const ends = (xs: number[]) => {
      const kept = new Set(xs)
      let sum = 0
      for (let j = 0; j < m; j++) if (!kept.has(j)) sum += Math.min(j, m - 1 - j)
      return sum
    }
    const [x, y] = [back(true), back(false)]
    const better = ends(y) < ends(x) || (ends(y) === ends(x) && moves(y) < moves(x))
    ;(better ? y : x).forEach((j, i) => ([to[i], same[i]] = [j, j >= 0]))
  } else
    for (let i = 0, j = 0; i < n && j < m; i++) {
      const k = b.indexOf(a[i], j)
      if (k >= 0) [to[i], same[i], j] = [k, true, k + 1]
    }
  // the rules left between two kept ones, paired in turn
  for (let i = 0, j = 0; i < n; ) {
    if (to[i] >= 0) {
      j = to[i++] + 1
      continue
    }
    let stop = i
    while (stop < n && to[stop] < 0) stop++
    const end = stop < n ? to[stop] : m
    for (; i < stop; i++) if (j < end) to[i] = j++
  }
  return { to, same }
}

/** A carried rule's place among the rules written, `p` of the workbook's before it in the file: after the last of those still in their place (not `late`), else before the first after it, else at `p`. */
function place(p: number, to: number[], late: boolean[], count: number): number {
  for (let k = p - 1; k >= 0; k--) if (to[k] >= 0 && !late[k]) return to[k] + 0.5
  for (let k = p; k < to.length; k++) if (to[k] >= 0 && !late[k]) return to[k] - 0.5
  return Math.min(p, count) - 0.5
}

export interface KeptRulesInput {
  xml: string
  sheet: string
  sources: RuleSource[]
  /** The edits of the sheet's part. */
  edits: Edit[]
  /** Rows or columns were deleted on the sheet, and none inserted. */
  deleted: boolean
  r0: ConditionalFormat[]
  r1: ConditionalFormat[]
  /** The workbook's rules written with these priorities. */
  write(priority: (i: number) => number): ConditionalWritten
}

/** A carried rule: its text with its moved formulas, its moved ranges, its place among the workbook's rules, Excel 2010's part of a data bar. */
interface Carried {
  s: RuleSource
  text: string
  ranges: string[]
  at: number
  partner?: Carried
}

const BAR_ID = /\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}/g

/**
 * The sheet's conditional formats written again: the workbook's (`r1`) with the file's rules the model does not have
 * among them, as the edits of their carriers moved them; a rule the reader read without a part of its look keeps the
 * file's differential format where the workbook's look of it did not change.
 */
export function keptRules(x: KeptRulesInput): { main: string; ext: string } {
  const { xml, sheet, sources, r0, r1 } = x
  const refuse = (why: string) => new KeptRefusal(`the conditional formats of '${sheet}' changed, and ${why}`, 'rules')
  if (sources.some((s) => s.pivot)) throw refuse("the file has a pivot table's rules, which the pivot table names by priority")
  for (const s of sources) {
    const lost = s.lost.find((k) => !KEPT_LOSSES.has(k))
    if (s.index !== undefined && lost) throw refuse(`the workbook read a rule over ${s.block.sqref} without what it holds (${lost})`)
  }
  const { to, same } = align(r0, r1)
  const cells = (t: string) => (sqrefRects(t) ?? []).reduce((n, r) => n + cellCount(r), 0)
  /** A rule of the file as carried, its edits made; undefined for one whose cells were all deleted. */
  const carry = (s: RuleSource): (Carried & { was: string }) | undefined => {
    if (s.x14 && [...xml.slice(s.start, s.end).matchAll(/<\/?([\w.-]+):/g)].some((m) => m[1] !== 'x14' && m[1] !== 'xm')) throw refuse(`a rule of Excel 2010 over ${s.block.sqref} is in prefixes of its own`)
    const own = x.edits.filter((e) => e.start >= s.start && e.end <= s.end)
    if (own.some((e) => e.start === s.start && e.end === s.end)) return undefined
    const span = sqrefAt(xml, s)
    const was = span ? readText(xml.slice(...span)) : s.block.sqref
    const moved = span && x.edits.find((e) => e.start === span[0] && e.end === span[1])
    const text = edited(xml.slice(s.start, s.end), own.map((e) => ({ ...e, start: e.start - s.start, end: e.end - s.start })))
    return { s, text, was, ranges: (moved ? readText(moved.text) : was).split(/\s+/).filter(Boolean), at: 0 }
  }
  const area = (ranges: string[]) => cells(ranges.join(' '))
  // the workbook's rules Excel put after the others (`shiftConditional`): after deleted lines as `lateAfterDeletion` says, else those over the statistics of their cells that lost some
  const late = r0.map((r, i) => to[i] >= 0 && (x.deleted ? lateAfterDeletion(r, r1[to[i]], sheet, new Map()) : GRADED.has(r.type) && area(r1[to[i]].ranges) < area(r.ranges)))
  /** A carried rule that goes after the others: before the first of those after it in the file. */
  const lastAt = (p: number) => {
    for (let k = p; k < r0.length; k++) if (late[k]) return to[k] - 0.5
    return Infinity
  }
  const carried: Carried[] = []
  let p = 0
  for (const s of sources) {
    if (s.joined) continue
    if (s.index !== undefined) {
      p++
      continue
    }
    const text = xml.slice(s.start, s.end)
    const c = carry(s)
    if (!c) continue
    const others = s.id === undefined ? [] : sources.filter((o) => o !== s && o.id === s.id)
    if (others.some((o) => o.index !== undefined)) throw refuse(`the workbook read the Excel 2010 part of a data bar over ${s.block.sqref} without its rule`)
    const part = others.find((o) => o.joined && o.x14)
    if (part) c.partner = carry(part)
    // a rule goes after the others as the workbook's would (its formulas without relative references, or with its first cell kept, compare as written)
    const type = attributes(text.slice(0, tagEnd(text, 0))).type ?? ''
    const cut = cells(c.ranges.join(' ')) < cells(c.was)
    const d = { cut, stayed: c.ranges.join(' ') === c.was, changed: c.text !== text, lost: lostRefs([c.text]) > lostRefs([text]) }
    c.at = (x.deleted ? lateBy(type, d) : GRADED.has(type) && cut) ? lastAt(p) : place(p, to, late, r1.length)
    carried.push(c)
  }
  // the rules in their order: the workbook's at their places, the carried ones between them
  const order = [...r1.map((_, j) => ({ at: j, j })), ...carried.map((c) => ({ at: c.at, c }))].sort((a, b) => a.at - b.at)
  const priority = new Map<number, number>()
  order.forEach((o, k) => 'j' in o && priority.set(o.j, k + 1))
  const written = x.write((i) => priority.get(i)!)
  const each = written.each.map((w) => w && { ...w })
  keepLooks(xml, sources, r0, r1, to, same, each)
  // a carried data bar's id that the workbook's rules use is renamed
  const taken = new Set(each.flatMap((w) => [...`${w?.main ?? ''}${w?.ext ?? ''}`.matchAll(BAR_ID)].map((m) => m[0].toUpperCase())))
  let fresh = 0
  const freshId = () => {
    let id: string
    do id = `{FFFFFFFF-0000-4000-8000-${(++fresh).toString(16).toUpperCase().padStart(12, '0')}}`
    while (taken.has(id))
    return id
  }
  let main = ''
  let ext = ''
  for (const [k, o] of order.entries()) {
    if ('j' in o) {
      main += each[o.j]?.main ?? ''
      ext += each[o.j]?.ext ?? ''
      continue
    }
    const { c } = o
    let text = c.text.slice(0, tagEnd(c.text, 0))
    text = setAttrs(text, { priority: String(k + 1) }) + c.text.slice(text.length)
    let partner = c.partner?.text
    const id = c.s.id
    if (id !== undefined && taken.has(id.toUpperCase())) {
      const next = freshId()
      text = text.replace(id, next)
      partner = partner?.replace(id, next)
    }
    if (c.s.x14) ext += block14(c.ranges, text)
    else main += block(c.ranges, text)
    if (c.partner && partner) ext += block14(c.partner.ranges, partner)
  }
  return { main, ext: x14Ext(ext) }
}

/**
 * The file's differential format of each rule the reader read without a part of its look (`condStyleLost`), in the
 * workbook's writing of the same rule: its `dxfId`, or Excel 2010's own `x14:dxf`.
 */
function keepLooks(xml: string, sources: RuleSource[], r0: ConditionalFormat[], r1: ConditionalFormat[], to: number[], same: boolean[], each: ({ main: string; ext: string } | undefined)[]): void {
  const used = new Set(to.filter((j, i) => same[i] && j >= 0))
  for (const s of sources) {
    if (s.index === undefined || s.joined || !s.lost.includes('condStyleLost')) continue
    let j = same[s.index] ? to[s.index] : -1
    if (j < 0) {
      // a rule that moved among the others
      j = r1.findIndex((r, k) => !used.has(k) && shape(r) === shape(r0[s.index!]))
      if (j < 0) continue
      used.add(j)
    }
    const w = each[j]
    const text = xml.slice(s.start, s.end)
    if (!w) continue
    if (!s.x14) {
      const dxfId = attributes(text.slice(0, tagEnd(text, 0))).dxfId
      if (dxfId !== undefined) w.main = w.main.replace(/<cfRule\b[^>]*>/, (tag) => setAttrs(tag, { dxfId }))
      continue
    }
    const own = placed(text, 'dxf')[0]
    const dxf = own && text.slice(own.start, own.end)
    if (!dxf || !w.ext || [...dxf.matchAll(/<\/?([\w.-]+):/g)].some((m) => m[1] !== 'x14')) continue
    w.ext = /<x14:dxf>[\s\S]*?<\/x14:dxf>/.test(w.ext) ? w.ext.replace(/<x14:dxf>[\s\S]*?<\/x14:dxf>/, () => dxf) : w.ext.replace('</x14:cfRule>', `${dxf}</x14:cfRule>`)
  }
}
