/**
 * Text typed into a session's cell (`Session.setCellText`): read as `parseEntry` reads it in the book's locale, date
 * system and clock, by the number format the cell shows, its own or else its row's or column's. The cell keeps the rest
 * of its input, and a format it shows from its line is not written into it. Beside the cell's input: why Excel would not
 * take the text, the width its column takes as Excel widens it for a typed number, and the rule of data validation the
 * value breaks.
 */
import type { CellInput, CellReplace, CellTextRefusal, DataValidation, EntryCheck, Refusal, ValidationAlert } from '../contract.js'
import type { EngineState } from '../engine.js'
import { parseEntry } from '../entry/index.js'
import { entryWidth } from '../layout/fit.js'
import { formatOf, sheetLines } from '../layout/lines.js'
import { parseMerge } from '../layout/merges.js'
import { hasCell } from '../structure/area.js'
import { say } from '../locale/messages.js'
import { parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { arrayPart } from '../spill/fixed.js'
import { brokenRule } from '../validation/session.js'
import { cellValue } from './looks.js'

/** The fields an entry replaces: the content and the number format. */
const TYPED = ['value', 'formula', 'array', 'numberFormat'] as const

/** The input `own` of the cell after `text` is typed into it at the moment `now` (the year of a date typed without one), or why Excel stops the entry. */
export function cellText(s: EngineState, i: number, r: number, c: number, own: CellInput | undefined, text: string, now: string): { ok: true; input: CellInput | null } | Refusal<CellTextRefusal> {
  const lines = sheetLines(s.input.sheets[i], s)
  const inherited = formatOf(undefined, lines.row(r), lines.column(c))
  const shown = own?.numberFormat ?? inherited
  const p = parseEntry(text, { locale: s.input.locale, dateSystem: s.input.dateSystem, now, cell: shown === undefined ? null : { numberFormat: shown } })
  if (!p.ok) return p as Refusal<CellTextRefusal>
  const typed = p.input ?? {}
  const out: Record<string, unknown> = { ...own }
  for (const k of TYPED) delete out[k]
  if (typed.formula !== undefined) out.formula = typed.formula
  else if (typed.value !== undefined) out.value = typed.value
  // a cell in General keeps its own General; a format its row or column gives it stays theirs
  const format = typed.numberFormat ?? own?.numberFormat
  if (format !== undefined && (own?.numberFormat !== undefined || format !== inherited)) out.numberFormat = format
  for (const k in out) if (out[k] !== undefined) return { ok: true, input: out as CellInput }
  return { ok: true, input: null }
}

/** Why Excel would not take an entry into the cell: it changes part of an array formula, or puts a formula into a table's header cell; null when it takes it. */
export function textRefusal(s: EngineState, edit: CellReplace): Refusal<CellTextRefusal> | null {
  const part = arrayPart(s, [edit])
  if (part) return { ok: false, code: 'array', message: say(s.input.locale, { key: 'arrayPart', ...part }) }
  if (edit.input?.formula === undefined) return null
  const i = s.sheetIdx.get(sheetKey(edit.sheet))!
  const { row, col } = parseAddress(edit.cell)
  const table = s.tables.list.find((t) => t.headers && t.s === i && t.r1 === row && col >= t.c1 && col <= t.c2)
  return table ? { ok: false, code: 'header', message: say(s.input.locale, { key: 'headerFormula', sheet: s.sheetNames[i], cell: edit.cell, table: table.name }) } : null
}

/**
 * The width the cell's column takes for the value the cell holds now, as Excel widens a column of the default width, or
 * one it fitted, for a number typed into it (`entryWidth`); undefined when it keeps its width, and for a hidden column
 * or a cell merged across columns, which the measurements leave out.
 */
export function columnWidening(s: EngineState, i: number, r: number, c: number): number | undefined {
  const sh = s.input.sheets[i]
  const lines = sheetLines(sh, s)
  const column = lines.column(c)
  if (column?.hidden || column?.width === 0) return undefined
  if ((sh.merges ?? []).some((m) => {
    const x = parseMerge(m)!
    return x.c2 > x.c1 && hasCell(x, r, c)
  })) return undefined
  const { locale, dateSystem } = s.input
  return entryWidth(cellValue(s, i, r, c), { numberFormat: lines.format(r, c), locale, dateSystem, width: column?.width, bestFit: column?.bestFit })
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
