/** Formula dependencies through shared column/row intervals, without expanding every referenced formula into an edge. */
import { spend } from '../exec/budget.js'
import { lowerBound } from '../sorted.js'
import type { Area } from '../spill/state.js'

interface Cell { r: number; id: number }
interface Rows { cells: Cell[]; nodes: Map<number, number> }
interface Column { c: number; rows: Rows }
interface Sheet { cols: Column[]; byCol: Map<number, Rows>; groups: Map<number, Rows> }
type FormulaCell = { s: number; r: number; c: number; id: number }
const rowOf = (x: Cell) => x.r
const colOf = (x: Column) => x.c
const DIRECT = 8

export class RangeDependencies {
  /** Formula ids keep their places; helper vertices are appended after them. Edges point toward readers. */
  readonly succ: number[][]
  private readonly sheets = new Map<number, Sheet>()

  constructor(formulas: readonly (FormulaCell | undefined)[]) {
    this.succ = Array.from(formulas, () => [])
    for (const f of formulas) {
      if (!f) continue
      let sh = this.sheets.get(f.s)
      if (!sh) this.sheets.set(f.s, sh = { cols: [], byCol: new Map(), groups: new Map() })
      let rows = sh.byCol.get(f.c)
      if (!rows) {
        sh.byCol.set(f.c, rows = { cells: [], nodes: new Map() })
        sh.cols.push({ c: f.c, rows })
      }
      rows.cells.push({ r: f.r, id: f.id })
    }
    for (const sh of this.sheets.values()) {
      sh.cols.sort((a, b) => a.c - b.c)
      for (const rows of sh.byCol.values()) rows.cells.sort((a, b) => a.r - b.r)
    }
  }

  at(s: number, r: number, c: number): number | undefined {
    const cells = this.sheets.get(s)?.byCol.get(c)?.cells
    const i = cells ? lowerBound(cells, r, rowOf) : -1
    return cells?.[i]?.r === r ? cells[i].id : undefined
  }

  add(area: Area, reader: number): void {
    const sh = this.sheets.get(area.s)
    if (!sh) return
    if (area.c1 === area.c2) {
      const rows = sh.byCol.get(area.c1)
      if (rows) this.linkRows(rows, area, reader)
      return
    }
    const first = lowerBound(sh.cols, area.c1, colOf), last = lowerBound(sh.cols, area.c2 + 1, colOf)
    const visit = (p: number, lo: number, hi: number) => {
      if (first >= hi || last <= lo) return
      if (first <= lo && hi <= last) {
        let rows = hi - lo === 1 ? sh.cols[lo].rows : sh.groups.get(p)
        if (!rows) {
          const cells: Cell[] = []
          for (let c = lo; c < hi; c++) {
            spend(sh.cols[c].rows.cells.length)
            for (const cell of sh.cols[c].rows.cells) cells.push(cell)
          }
          cells.sort((a, b) => { spend(1); return a.r - b.r || a.id - b.id })
          sh.groups.set(p, rows = { cells, nodes: new Map() })
        }
        this.linkRows(rows, area, reader)
        return
      }
      const mid = (lo + hi) >>> 1
      visit(p * 2, lo, mid); visit(p * 2 + 1, mid, hi)
    }
    if (first < last) visit(1, 0, sh.cols.length)
  }

  private linkRows(rows: Rows, area: Area, reader: number): void {
    const cells = rows.cells
    const first = lowerBound(cells, area.r1, rowOf), last = lowerBound(cells, area.r2 + 1, rowOf)
    const visit = (p: number, lo: number, hi: number) => {
      if (first >= hi || last <= lo) return
      if (hi - lo <= DIRECT) {
        for (let i = Math.max(first, lo); i < Math.min(last, hi); i++) this.succ[cells[i].id].push(reader)
      } else if (first <= lo && hi <= last) this.succ[this.vertex(rows, p, lo, hi)].push(reader)
      else {
        const mid = (lo + hi) >>> 1
        visit(p * 2, lo, mid); visit(p * 2 + 1, mid, hi)
      }
    }
    if (first < last) visit(1, 0, cells.length)
  }

  private vertex(rows: Rows, p: number, lo: number, hi: number): number {
    if (hi - lo === 1) return rows.cells[lo].id
    const hit = rows.nodes.get(p)
    if (hit !== undefined) return hit
    const id = this.succ.length
    rows.nodes.set(p, id)
    this.succ.push([])
    spend(1)
    if (hi - lo <= DIRECT) {
      for (let i = lo; i < hi; i++) this.succ[rows.cells[i].id].push(id)
    } else {
      const mid = (lo + hi) >>> 1
      this.succ[this.vertex(rows, p * 2, lo, mid)].push(id)
      this.succ[this.vertex(rows, p * 2 + 1, mid, hi)].push(id)
    }
    return id
  }
}
