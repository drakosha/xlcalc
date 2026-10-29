/**
 * Text typed into a session's cell (`Session.setCellText`), or pasted into it: read as `parseEntry` reads it in the book's
 * locale, date system and clock (pasted, as `readEntry` reads it), by the number format the cell shows, its own or else
 * its row's or column's. The cell keeps the rest of its input, and a format it shows from its line is not written into it. Beside the cell's input: why Excel would not
 * take the text, the width its column takes as Excel widens it for a typed number, and the rule of data validation the
 * value breaks.
 */
import type { CellInput, CellRef, CellReplace, CellTextRefusal, DataValidation, EntryCheck, EntryResult, Refusal, ValidationAlert } from '../contract.js'
import { sheetIndexOf, type EngineState } from '../engine.js'
import { readEntry, type EntryWay } from '../entry/index.js'
import type { FormatAt } from '../entry/formula.js'
import { entryWidth } from '../layout/fit.js'
import { formatOf, sheetLines } from '../layout/lines.js'
import { mergedOf } from '../layout/merges.js'
import { hasCell } from '../structure/area.js'
import { say } from '../locale/messages.js'
import { parseAddress } from '../model.js'
import { arrayPart } from '../spill/fixed.js'
import { brokenRule } from '../validation/session.js'
import { CONTENT, someField } from './cells.js'
import { cellValue } from './looks.js'
import { turnOf } from '../layout/turned.js'

/** The fields an entry replaces: the content and the number format. */
const TYPED = [...CONTENT, 'numberFormat'] as const

/** The number format the cell's row or column gives it. */
function lineFormat(s: EngineState, i: number, r: number, c: number) {
  const lines = sheetLines(s.input.sheets[i], s)
  return formatOf(undefined, lines.row(r), lines.column(c))
}

/** The input a cell got from an entry earlier in the same call (null: none left), by sheet index and zero-based row and column; undefined where it got none. */
export type EarlierInput = (i: number, r: number, c: number) => CellInput | null | undefined

/** The format a cell of the book shows, by its sheet's name (null: sheet `i`), as a formula typed into sheet `i` refers to it, after the call's earlier entries; none on a sheet the book lacks. */
const formatIn =
  (s: EngineState, i: number, earlier?: EarlierInput): FormatAt =>
  (sheet, r, c) => {
    const k = sheet === null ? i : sheetIndexOf(s, sheet)
    if (k === undefined) return undefined
    const own = earlier?.(k, r, c)
    return sheetLines(s.input.sheets[k], s).format(r, c, own === null ? {} : own)
  }

/** What `text` typed, edited in place or pasted (`way`) into the cell at the moment `now` reads as (`readEntry`), by the format the cell shows: `own`'s, else its row's or column's; its references by the formats of the book after `earlier`. */
export function readText(s: EngineState, i: number, r: number, c: number, own: CellInput | undefined, text: string, now: string, way: EntryWay = 'typed', earlier?: EarlierInput): EntryResult {
  const shown = own?.numberFormat ?? lineFormat(s, i, r, c)
  return readEntry(text, { locale: s.input.locale, dateSystem: s.input.dateSystem, now, cell: shown === undefined ? null : { numberFormat: shown }, externalBooks: s.input.externalBooks ?? [] }, way, formatIn(s, i, earlier))
}

/** The input `own` of the cell with an entry's content and format (`typed`, `readText`'s input) in place of its own; a format its row or column gives it stays theirs. */
export function withTyped(s: EngineState, i: number, r: number, c: number, own: CellInput | undefined, typed: CellInput | null): CellInput | null {
  const inherited = lineFormat(s, i, r, c)
  const out: Record<string, unknown> = { ...own }
  for (const k of TYPED) delete out[k]
  if (typed?.formula !== undefined) out.formula = typed.formula
  else if (typed?.value !== undefined) out.value = typed.value
  // a cell in General keeps its own General; a format its row or column gives it stays theirs
  const format = typed?.numberFormat ?? own?.numberFormat
  if (format !== undefined && (own?.numberFormat !== undefined || format !== inherited)) out.numberFormat = format
  return someField<CellInput>(out)
}

/** The input `own` of the cell after `text` is typed, edited in place or pasted (`way`) into it at the moment `now` (the year of a date typed without one), or why Excel stops the entry. */
export function cellText(s: EngineState, i: number, r: number, c: number, own: CellInput | undefined, text: string, now: string, way: EntryWay = 'typed', earlier?: EarlierInput): { ok: true; input: CellInput | null } | Refusal<CellTextRefusal> {
  const p = readText(s, i, r, c, own, text, now, way, earlier)
  return p.ok ? { ok: true, input: withTyped(s, i, r, c, own, p.input) } : (p as Refusal<CellTextRefusal>)
}

/** Why Excel would not take the edits: they change part of an array formula; null when it takes them. */
export function arrayRefusal(s: EngineState, edits: CellReplace[]): (Refusal<'array'> & CellRef) | null {
  const part = arrayPart(s, edits)
  return part && { ok: false, code: 'array', message: say(s.input.locale, { key: 'arrayPart', ...part }), sheet: part.sheet, cell: part.cell }
}

/** Why Excel would not take an entry into the cell: it puts a formula into a table's header cell; null when it takes it. */
export function headerRefusal(s: EngineState, edit: CellReplace): Refusal<'header'> | null {
  if (edit.input?.formula === undefined) return null
  const i = sheetIndexOf(s, edit.sheet)!
  const { row, col } = parseAddress(edit.cell)
  const table = s.tables.list.find((t) => t.headers && t.s === i && t.r1 === row && col >= t.c1 && col <= t.c2)
  return table ? { ok: false, code: 'header', message: say(s.input.locale, { key: 'headerFormula', sheet: s.sheetNames[i], cell: edit.cell, table: table.name }) } : null
}

/** Why Excel would not take an entry into the cell alone: it changes part of an array formula, or puts a formula into a table's header cell; null when it takes it. */
export function textRefusal(s: EngineState, edit: CellReplace): Refusal<CellTextRefusal> | null {
  const part = arrayRefusal(s, [edit])
  return part ? { ok: false, code: 'array', message: part.message } : headerRefusal(s, edit)
}

/**
 * The width the cell's column takes for the value the cell holds now, as Excel widens a column of the default width, or
 * one it fitted, for a number typed into it (`entryWidth`); undefined when it keeps its width, and for a hidden column
 * or a cell merged across columns, which the measurements leave out.
 */
export function columnWidening(s: EngineState, i: number, r: number, c: number, memo?: Map<string, number | undefined>): number | undefined {
  const sh = s.input.sheets[i]
  const lines = sheetLines(sh, s)
  const column = lines.column(c)
  if (column?.hidden || column?.width === 0) return undefined
  if (mergedOf(sh, s).rects.some((x) => x.c2 > x.c1 && hasCell(x, r, c))) return undefined
  // a turned cell lays its text across the column: Excel widens nothing for it (`test/excel/fit/rotation.json`, `typed`)
  const style = lines.look(r, c).style
  if (style !== undefined && turnOf(s.input.styles?.[style]) !== undefined) return undefined
  const { locale, dateSystem } = s.input
  const value = cellValue(s, i, r, c)
  const numberFormat = lines.format(r, c)
  // a column's cells of one value and format take one width (a fill of a constant): measured once in a call that passes `memo`
  const k = memo && JSON.stringify([i, c, numberFormat ?? null, value])
  if (k !== undefined && memo!.has(k)) return memo!.get(k)
  const width = entryWidth(value, { numberFormat, locale, dateSystem, width: column?.width, bestFit: column?.bestFit })
  if (k !== undefined) memo!.set(k, width)
  return width
}

/** A rule's alert: its style (stop by default) and its message, the rule's own or else the library's in the book's locale. */
const alertOf = (s: EngineState, rule: DataValidation) => ({ rule, severity: rule.errorStyle ?? 'stop', message: rule.error ?? say(s.input.locale, { key: 'validationError' }) })

/** The alert Excel shows for the value the cell holds now against its rule of data validation; none when it meets it, or when the rule shows no alert. */
export function validationAlert(s: EngineState, i: number, r: number, c: number): ValidationAlert | undefined {
  const rule = brokenRule(s, i, r, c)
  return rule && rule.showError !== false ? { valid: false, ...alertOf(s, rule) } : undefined
}

/** Whether the value the cell holds now meets its rule (`Session.validateEntry`). */
export function entryCheck(s: EngineState, i: number, r: number, c: number): EntryCheck {
  const rule = brokenRule(s, i, r, c)
  return rule ? { ok: false, code: 'validation', ...alertOf(s, rule) } : { ok: true }
}
