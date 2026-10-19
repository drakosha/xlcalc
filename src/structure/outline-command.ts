/**
 * Excel's outline commands on a sheet's rows and columns, without the engine: the outline operations of `editStructure`
 * (Group and Ungroup of rows or columns, Show and Hide Detail at a line, the level buttons) read into the commands of
 * `layout/outline.ts` and applied to one axis each, or refused as Excel refuses them. `structure/outline.ts` applies
 * them to a sheet of a book, with a message in the book's locale; `outlineEdit` of the `./a1` subpath to a sheet alone.
 */
import type { SheetInput, StructureOp } from '../contract.js'
import { InputError } from '../input-error.js'
import { integer } from '../input/basic.js'
import { outlineCommand, type OutlineCommand, type OutlineRefusal } from '../layout/outline.js'
import { checkSpan, spanKey } from '../layout/spans.js'
import type { OutlineMessageKey } from '../locale/outline-messages.js'
import type { Lines } from '../locale/messages.js'
import { colToIndex, MAX_COLS, MAX_ROWS } from '../model.js'
import type { Axis } from './area.js'

export type OutlineOp = Extract<StructureOp, { op: 'group' | 'ungroup' | 'showDetail' | 'hideDetail' | 'showLevels' }>

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
      if (axis === 'r') return [['r', { kind, at: integer(o.row, 'op.row', 1, MAX_ROWS) - 1 }]]
      const at = typeof o.column === 'string' && /^[A-Z]{1,3}$/.test(o.column) ? colToIndex(o.column) : MAX_COLS
      if (at >= MAX_COLS) throw new InputError('op.column', `expected column letters from A to XFD, got ${JSON.stringify(o.column)}`)
      return [['c', { kind, at }]]
    }
    case 'showLevels': {
      if (o.rows === undefined && o.columns === undefined) throw new InputError('op', 'expected rows, columns or both')
      const out: [Axis, OutlineCommand][] = []
      if (o.rows !== undefined) out.push(['r', { kind: 'levels', level: integer(o.rows, 'op.rows', 1, 8) }])
      if (o.columns !== undefined) out.push(['c', { kind: 'levels', level: integer(o.columns, 'op.columns', 1, 8) }])
      return out
    }
  }
}

/** The lines a command names, for its refusal's message: `5:7`, `C`. */
const named = (axis: Axis, c: OutlineCommand) => ('at' in c ? spanKey(c.at, c.at, axis) : 'lo' in c ? spanKey(c.lo, c.hi, axis) : '')

export type OutlineAxes = Pick<SheetInput, 'rows' | 'columns' | 'outline'>

/** Why Excel refuses an outline command: its code, and its message's key and parameters. */
export interface OutlineRefused {
  ok: false
  code: OutlineRefusal
  message: { key: OutlineMessageKey; lines: Lines; span: string }
}

/**
 * The sheet after the operation (the same object when nothing changes; an axis left with no lines is dropped), or why
 * Excel refuses it. The operation's fields but `op` and `sheet` are checked here.
 */
export function applyOutline<S extends OutlineAxes>(sheet: S, o: Record<string, unknown>): { ok: true; sheet: S } | OutlineRefused {
  let s = sheet
  for (const [axis, c] of commands(o)) {
    const key = axis === 'r' ? 'rows' : 'columns'
    const after = (axis === 'r' ? s.outline?.summaryBelow : s.outline?.summaryRight) !== false
    const r = outlineCommand(s[key], axis, after, c)
    if (typeof r === 'string') {
      const k = r === 'outline_deep' ? 'outlineDeep' : r === 'detail' ? (c.kind === 'show' ? 'detailShown' : 'detailHidden') : c.kind === 'ungroup' ? 'notGrouped' : 'noSummary'
      return { ok: false, code: r, message: { key: k, lines: key, span: named(axis, c) } }
    }
    if (r !== s[key]) {
      const { [key]: _, ...rest } = s
      s = (r && Object.keys(r).length ? { ...s, [key]: r } : rest) as S
    }
  }
  return { ok: true, sheet: s }
}
