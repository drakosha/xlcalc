/**
 * Excel's outline commands as structure operations (`editStructure`): Group and Ungroup of rows or columns, Show and Hide
 * Detail at a line, the level buttons; each changes one axis of one sheet as `layout/outline.ts` says, or is refused as
 * Excel refuses it.
 */
import type { BookInput, SheetInput, StructureOp } from '../contract.js'
import { checkSpan, InputError } from '../input.js'
import { say } from '../locale/messages.js'
import { outlineCommand, type OutlineCommand } from '../layout/outline.js'
import { spanKey } from '../layout/spans.js'
import { colToIndex, MAX_COLS, MAX_ROWS } from '../model.js'
import type { Axis } from './area.js'
import type { Applied } from './geometry.js'

type OutlineOp = Extract<StructureOp, { op: 'group' | 'ungroup' | 'showDetail' | 'hideDetail' | 'showLevels' }>

/** The fields each outline operation takes beside `op`. */
export const OUTLINE_FIELDS: Record<OutlineOp['op'], string[]> = {
  group: ['sheet', 'rows', 'columns'],
  ungroup: ['sheet', 'rows', 'columns'],
  showDetail: ['sheet', 'row', 'column'],
  hideDetail: ['sheet', 'row', 'column'],
  showLevels: ['sheet', 'rows', 'columns'],
}

/** One of two fields, the other left out. */
function oneOf(o: Record<string, unknown>, a: string, b: string): Axis {
  if ((o[a] === undefined) === (o[b] === undefined)) throw new InputError('op', `expected one of ${a} and ${b}`)
  return o[a] !== undefined ? 'r' : 'c'
}

function level(v: unknown, path: string): number {
  if (typeof v !== 'number' || !Number.isInteger(v) || v < 1 || v > 8) throw new InputError(path, `expected an integer from 1 to 8, got ${JSON.stringify(v)}`)
  return v
}

/** The commands of the operation, by axis. */
function commands(o: Record<string, unknown>): [Axis, OutlineCommand][] {
  switch (o.op as OutlineOp['op']) {
    case 'group':
    case 'ungroup': {
      const axis = oneOf(o, 'rows', 'columns')
      const key = axis === 'r' ? 'rows' : 'columns'
      const [lo, hi] = checkSpan(o[key], `op.${key}`, axis)
      return [[axis, { kind: o.op as 'group' | 'ungroup', lo, hi }]]
    }
    case 'showDetail':
    case 'hideDetail': {
      const axis = oneOf(o, 'row', 'column')
      const kind = o.op === 'showDetail' ? 'show' : 'hide'
      if (axis === 'r') {
        if (typeof o.row !== 'number' || !Number.isInteger(o.row) || o.row < 1 || o.row > MAX_ROWS) throw new InputError('op.row', `expected an integer from 1 to ${MAX_ROWS}, got ${JSON.stringify(o.row)}`)
        return [['r', { kind, at: o.row - 1 }]]
      }
      const at = typeof o.column === 'string' && /^[A-Z]{1,3}$/.test(o.column) ? colToIndex(o.column) : MAX_COLS
      if (at >= MAX_COLS) throw new InputError('op.column', `expected column letters from A to XFD, got ${JSON.stringify(o.column)}`)
      return [['c', { kind, at }]]
    }
    case 'showLevels': {
      if (o.rows === undefined && o.columns === undefined) throw new InputError('op', 'expected rows, columns or both')
      const out: [Axis, OutlineCommand][] = []
      if (o.rows !== undefined) out.push(['r', { kind: 'levels', level: level(o.rows, 'op.rows') }])
      if (o.columns !== undefined) out.push(['c', { kind: 'levels', level: level(o.columns, 'op.columns') }])
      return out
    }
  }
}

/** The lines a command names, for its refusal's message: `5:7`, `C`. */
const named = (axis: Axis, c: OutlineCommand) => ('at' in c ? spanKey(c.at, c.at, axis) : 'lo' in c ? spanKey(c.lo, c.hi, axis) : '')

/** The operation on the sheet `sheet` of a book already checked, its other fields checked here. */
export function outlineOp(book: BookInput, sheet: string, o: Record<string, unknown>): Applied {
  const at = book.sheets.findIndex((s) => s.name === sheet)
  let s: SheetInput = book.sheets[at]
  for (const [axis, c] of commands(o)) {
    const key = axis === 'r' ? 'rows' : 'columns'
    const after = (axis === 'r' ? s.outline?.summaryBelow : s.outline?.summaryRight) !== false
    const r = outlineCommand(s[key], axis, after, c)
    if (typeof r === 'string') {
      const p = { lines: key, span: named(axis, c) } as const
      const message = say(book.locale, { key: r === 'outline_deep' ? 'outlineDeep' : r === 'detail' ? (c.kind === 'show' ? 'detailShown' : 'detailHidden') : c.kind === 'ungroup' ? 'notGrouped' : 'noSummary', ...p })
      return { ok: false, code: r, message }
    }
    if (r !== s[key]) {
      const { [key]: _, ...rest } = s
      s = r && Object.keys(r).length ? { ...s, [key]: r } : rest
    }
  }
  const sheets = book.sheets.slice()
  sheets[at] = s
  return { ok: true, book: s === book.sheets[at] ? book : { ...book, sheets }, geometry: [] }
}
