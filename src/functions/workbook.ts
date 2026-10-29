/**
 * Functions about references and the book, this engine's own: CELL, INFO, SHEET, SHEETS, FORMULATEXT and ISFORMULA.
 * They read what the book holds beside values through `Cx.book` (`book-facts.ts`): a cell's formula and look, its
 * sheets, the file it was opened from. INFO's facts of the system are constants, Excel for Windows' own.
 */
import type { Arg, Cx, FnDef } from '../args.js'
import { arr } from '../arrays/value.js'
import { indexToCol } from '../model.js'
import { quoteSheet } from '../syntax.js'
import { Areas, E, isErr, Rng, type V } from '../values.js'
import { formatCodes } from './cell-format.js'

/** The longest formula text FORMULATEXT gives: a longer one is #N/A. */
const LONGEST = 8192

/** A reference's top-left cell on one of the book's sheets, of its first area with `areas`, or the error a function answers for anything else. */
function topLeft(cx: Cx, x: Arg, areas = false): Rng | V {
  if (isErr(x)) return x
  const g = x instanceof Rng ? x : areas && x instanceof Areas ? x.list[0] : undefined
  return g && g.s < cx.book.sheets ? new Rng(g.s, g.r1, g.c1, g.r1, g.c1) : E.VALUE
}

/** A cell's address as CELL gives it: `$A$1`, on another sheet with the book's name (`[Book1]Sheet2!$A$1`). */
function address(cx: Cx, g: Rng, s: number): string {
  const a = `$${indexToCol(g.c1)}$${g.r1 + 1}`
  if (g.s === s) return a
  // quoted where the book's or the sheet's name has more than letters, digits, `_` and `.` (`[1.xlsx]R!$B$3`)
  const both = `[${cx.book.file?.name ?? 'Book1'}]${cx.book.sheetName(g.s)}`
  return `${/^\[[\p{L}\p{N}_.]+\][\p{L}\p{N}_.]+$/u.test(both) ? both : quoteSheet(both)}!${a}`
}

/** CELL's `prefix`: a text's alignment as the label prefix of old spreadsheets. */
const PREFIX: Record<string, string> = { general: "'", left: "'", right: '"', center: '^', fill: '\\' }

function cell(cx: Cx, s: number, r: number, c: number, info: string, ref?: Arg): Arg {
  let g: Rng | V
  if (ref === undefined) {
    // the active cell, which the graph learns as INDIRECT's reference
    const a = cx.book.activeCell()
    g = new Rng(a.s, a.r, a.c, a.r, a.c)
    cx.found(g)
  } else g = topLeft(cx, ref, true)
  if (!(g instanceof Rng)) return g
  const v = () => cx.scalar(g)
  switch (info.toLowerCase()) {
    case 'address':
      return address(cx, g, s)
    case 'col':
      return g.c1 + 1
    case 'row':
      return g.r1 + 1
    case 'contents':
      return v()
    case 'type': {
      const x = v()
      return x === null ? 'b' : typeof x === 'string' ? 'l' : 'v'
    }
    case 'filename': {
      const f = cx.book.file
      return f ? `${f.path ?? ''}[${f.name}]${cx.book.sheetName(g.s)}` : ''
    }
    case 'format':
      return formatCodes(cx.book.cell(g.s, g.r1, g.c1).format).format
    case 'color':
      return formatCodes(cx.book.cell(g.s, g.r1, g.c1).format).color
    case 'parentheses':
      return formatCodes(cx.book.cell(g.s, g.r1, g.c1).format).parentheses
    case 'prefix': {
      const x = v()
      return typeof x === 'string' ? (PREFIX[cx.book.cell(g.s, g.r1, g.c1).horizontal ?? 'general'] ?? '') : ''
    }
    case 'protect':
      return cx.book.cell(g.s, g.r1, g.c1).locked ? 1 : 0
    case 'width': {
      // Excel 365 answers the width and whether it is the sheet's default
      const x = cx.book.cell(g.s, g.r1, g.c1)
      return arr(1, 2, [Math.round(x.width), x.defaultWidth])
    }
    default:
      return E.VALUE
  }
}

/** INFO's answers: those of Excel for Windows, the book's own where it has them. */
function info(cx: Cx, type: string): V {
  switch (type.toLowerCase()) {
    case 'directory':
      return cx.book.file?.path ?? ''
    case 'numfile':
      return cx.book.sheets
    case 'origin': {
      const t = cx.book.topLeft()
      return `$A:$${indexToCol(t.c)}$${t.r + 1}`
    }
    case 'osversion':
      return 'Windows (64-bit) NT 10.00'
    case 'recalc':
      return cx.book.manual ? 'Manual' : 'Automatic'
    case 'release':
      return '16.0'
    case 'system':
      return 'pcdos'
    case 'memavail':
    case 'memused':
    case 'totmem':
      return E.NA
    default:
      return E.VALUE
  }
}

/** SHEET: the number of a reference's sheet, of a sheet by its name, or of the formula's own. */
function sheet(cx: Cx, s: number, _r: number, _c: number, x?: Arg): V {
  if (x === undefined) return s + 1
  if (isErr(x)) return x
  if (x instanceof Rng) return x.s < cx.book.sheets ? x.s + 1 : E.NA
  if (x instanceof Areas) return x.list[0].s < cx.book.sheets ? x.list[0].s + 1 : E.NA
  const v = cx.scalar(x)
  if (typeof v !== 'string') return isErr(v) ? v : E.NA
  const i = cx.book.sheetOf(v)
  return i === undefined ? E.NA : i + 1
}

/** SHEETS: the sheets a reference covers, or the book's. */
function sheets(cx: Cx, x?: Arg): V {
  if (x === undefined) return cx.book.sheets
  if (isErr(x)) return x
  if (x instanceof Rng) return 1
  if (x instanceof Areas) return x.cube ? x.list.length : 1
  return E.VALUE
}

export const functions: Record<string, FnDef> = {
  // a union's or a reference across sheets' first area
  CELL: { args: 's R?', cx: true, cell: true, dynamic: true, looks: true, array: true, areas: true, fn: (cx: Cx, s: number, r: number, c: number, type: string, ref?: Arg) => cell(cx, s, r, c, type, ref) },
  FORMULATEXT: {
    args: 'R',
    cx: true,
    fn: (cx: Cx, x: Arg) => {
      const g = topLeft(cx, x)
      if (!(g instanceof Rng)) return isErr(g) ? (g.code === 'VALUE' ? E.NA : g) : E.NA
      const text = cx.book.formula(g.s, g.r1, g.c1)
      return text === undefined || text.length > LONGEST ? E.NA : text
    },
  },
  INFO: { args: 's', cx: true, fn: info },
  ISFORMULA: {
    args: 'R',
    cx: true,
    fn: (cx: Cx, x: Arg) => {
      const g = topLeft(cx, x)
      return g instanceof Rng ? cx.book.hasFormula(g.s, g.r1, g.c1) : g
    },
  },
  SHEET: { args: 'R?', cx: true, cell: true, fn: sheet },
  SHEETS: { args: 'R?', cx: true, areas: true, fn: sheets },
}
