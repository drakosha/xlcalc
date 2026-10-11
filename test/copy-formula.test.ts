/**
 * `copyFormula(text, from, to, book?)`: the text of a copied formula. Against `editStructure`'s `copyRange` on every copy
 * of the structure sets and of the table sets: each formula of the source, at each place the copy pastes it, is the
 * text `copyFormula` gives (but where `copyRange` does what one formula cannot: rows a filter hid, part of an
 * old-kind array, a table copied whole with its references to the copy, a formula pasted on a table's header, which
 * becomes its text). And the rules case by case.
 */
import { describe, expect, it } from 'vitest'
import { buildTableCases as moreTables } from '../bench/coverage/tables/more.js'
import { buildTableCases as heldTables } from '../bench/coverage/tables/ops-held.js'
import { buildTableCases as opsTables } from '../bench/coverage/tables/ops.js'
import type { TableCase } from '../bench/coverage/tables/cases.js'
import { takenBook } from '../bench/coverage/stored.js'
import { copyFormula, InputError, createSession, editStructure, type BookInput, type StructureOp } from '../src/index.js'
import { hiddenRows } from '../src/layout/hidden.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { parseRange } from '../src/structure/area.js'
import { sheetKey } from '../src/syntax.js'
import { tablesOf } from '../src/tables/model.js'
import { STRUCT_SETS } from './structure-excel.js'

type Copy = Extract<StructureOp, { op: 'copyRange' }>

/** The cells a copy compares, each source formula at each place it lands: [source, target, formula there]. */
function compared(book: BookInput, op: Copy): [string, string, string, string][] | null {
  const r = editStructure(book, op)
  if (!r.ok) return null
  const sheet = (b: BookInput, name: string) => b.sheets.find((s) => sheetKey(s.name) === sheetKey(name))!
  const from = sheet(book, op.from.sheet)
  // a copy from a filtered sheet leaves out the rows a filter hid: no formula alone does that
  if (from.filtered && hiddenRows(from)) return []
  const src = parseRange(op.from.range)!
  const target = parseRange(op.to.range)!
  const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
  const tiled = (target.r2 - target.r1 + 1) % h === 0 && (target.c2 - target.c1 + 1) % w === 0
  const offsets: [number, number][] = []
  if (tiled) for (let row = target.r1; row <= target.r2; row += h) for (let col = target.c1; col <= target.c2; col += w) offsets.push([row - src.r1, col - src.c1])
  else offsets.push([target.r1 - src.r1, target.c1 - src.c1])
  // the tables the copy made: their cells refer to the copies
  const made = new Set(tablesOf(r.book).list.map((t) => t.name).filter((n) => !tablesOf(book).byName.has(n.toUpperCase())))
  const after = tablesOf(r.book)
  const toSheet = sheet(r.book, op.to.sheet)
  const out: [string, string, string, string][] = []
  for (const [addr, cell] of Object.entries(from.cells)) {
    const { row, col } = parseAddress(addr)
    if (cell.formula === undefined || cell.array !== undefined || row < src.r1 || row > src.r2 || col < src.c1 || col > src.c2) continue
    for (const [dr, dc] of offsets) {
      const dest = formatAddress(row + dr, col + dc)
      const s = r.book.sheets.indexOf(toSheet)
      const inside = (t: (typeof after.list)[number]) => t.s === s && row + dr >= t.r1 && row + dr <= t.r2 && col + dc >= t.c1 && col + dc <= t.c2
      if (after.list.some((t) => inside(t) && (made.has(t.name) || (t.headers && row + dr === t.r1)))) continue
      out.push([`${from.name}!${addr}`, `${toSheet.name}!${dest}`, cell.formula, toSheet.cells[dest]?.formula ?? '(none)'])
    }
  }
  return out
}

function check(label: string, book: BookInput, op: Copy): number {
  const cells = compared(book, op)
  if (!cells) return 0
  for (const [from, to, formula, pasted] of cells) {
    const [fs, fa] = from.split('!')
    const [ts, ta] = to.split('!')
    expect(copyFormula(formula, { sheet: fs, cell: fa }, { sheet: ts, cell: ta }, book), `${label}: ${formula} from ${from} to ${to}`).toBe(pasted)
  }
  return cells.length
}

describe('copyFormula against copyRange', () => {
  it('every copy of the structure sets', () => {
    let n = 0
    for (const set of STRUCT_SETS) for (const c of set.cases) if (c.op.op === 'copyRange' && !c.via) n += check(`${set.name} ${c.id}`, takenBook(set.books[c.book]), c.op)
    expect(n).toBeGreaterThan(100)
  })

  it('every copy of the table sets, in the book as the steps before it left it', () => {
    let n = 0
    const cases: TableCase[] = [...opsTables(), ...moreTables(), ...heldTables()]
    for (const c of cases) {
      const s = createSession(takenBook(c.book))
      for (const step of c.steps ?? []) {
        if ('op' in step) {
          if (step.op.op === 'copyRange') n += check(c.id, s.getBook(), step.op)
          s.editStructure(step.op)
        } else if ('edit' in step) {
          const v = step.edit.value
          if (v !== null && typeof v === 'object' && 'local' in v) break
          s.setCells([{ sheet: step.edit.sheet, cell: step.edit.cell, input: v === null ? null : typeof v === 'object' ? { formula: v.formula } : { value: v } }])
        }
      }
    }
    expect(n).toBeGreaterThan(3)
  })
})

describe('copyFormula', () => {
  const at = (cell: string, sheet = 'S') => ({ sheet, cell })

  it('relative parts move, absolute ones stay, past the edge is #REF!', () => {
    expect(copyFormula('A1+$B$2+B$3+$C4+SUM(D5:E6)', at('F10'), at('H13'))).toBe('C4+$B$2+D$3+$C7+SUM(F8:G9)')
    expect(copyFormula('A1+B:B+3:3', at('C3'), at('D5'))).toBe('B3+C:C+5:5')
    expect(copyFormula('A1*2', at('B2'), at('A1'))).toBe('#REF!*2')
    expect(copyFormula('SUM(A1:A3)', at('B3'), at('B1'))).toBe('SUM(#REF!)')
    expect(copyFormula("'My data'!A1+Other!$A$1+'Jan:Mar'!B2", at('C1'), at('C2', 'T'))).toBe("'My data'!A2+Other!$A$1+'Jan:Mar'!B3")
    expect(copyFormula('E2#*rate', at('F2'), at('F3'))).toBe('E3#*rate')
    // text that does not parse is pasted as written
    expect(copyFormula('SUM(A1', at('B1'), at('B2'))).toBe('SUM(A1')
  })

  it('a formula copied out of its table names the table, as Excel pastes it; inside the table it stays', () => {
    const book: BookInput = {
      sheets: [
        {
          name: 'S',
          cells: { A1: { value: 'Item' }, B1: { value: 'Price' }, C1: { value: 'Twice' }, A2: { value: 'a' }, B2: { value: 3 }, C2: { formula: '[@Price]*2' }, A3: { value: 'b' }, B3: { value: 4 }, C3: { formula: '[@Price]*2' } },
          tables: [{ name: 'Sales', range: 'A1:C3', columns: [{ name: 'Item' }, { name: 'Price' }, { name: 'Twice', formula: '[@Price]*2' }] }],
        },
        { name: 'T', cells: {} },
      ],
    }
    expect(copyFormula('[@Price]*2', at('C2'), at('C3'), book)).toBe('[@Price]*2')
    expect(copyFormula('[@Price]*2', at('C2'), at('E2'), book)).toBe(copyRangeAt(book, 'C2', 'E2'))
    expect(copyFormula('[@Price]*2+B2', at('C2'), at('E2', 'T'), book)).toBe('Sales[[#This Row],[Price]]*2+D2')
    expect(copyFormula('[@Price]*2', at('C2'), at('E2'))).toBe('[@Price]*2')
  })

  it('cells and sheets are checked as input', () => {
    expect(() => copyFormula('A1', at('A0'), at('B1'))).toThrow(InputError)
    expect(() => copyFormula('A1', at('A1'), at('B1', 'Nope'), { sheets: [{ name: 'S' }] })).toThrow('to.sheet: unknown sheet "Nope"')
    expect(copyFormula('A1', at('A1', 's'), at('B1', 's'), { sheets: [{ name: 'S' }] })).toBe('B1')
  })
})

/** What copyRange pastes in `to` from `from` on sheet S. */
function copyRangeAt(book: BookInput, from: string, to: string): string | undefined {
  const r = editStructure(book, { op: 'copyRange', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to } })
  return r.ok ? r.book.sheets[0].cells[to]?.formula : undefined
}
