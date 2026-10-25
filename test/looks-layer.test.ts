/**
 * The layer of a sheet's looks (`src/looks/`) against the cells painted one by one: a list of rectangles read (each over
 * those before it, a rectangle without fields clearing), written in its one form and read back, cells written over it,
 * a book's entries of a look alone taken into it. Seeded random lists over a small corner of the sheet, and the forms of
 * whole rows and columns.
 */
import { describe, expect, it } from 'vitest'
import type { LookRange } from '../src/contract.js'
import { cellEdits, edited, NONE, sameLook, type Layer } from '../src/looks/layer.js'
import { carried, differing, shifted } from '../src/looks/move.js'
import { readRanges, writeRanges } from '../src/looks/ranges.js'
import { looksTaken } from '../src/looks/sheet.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../src/model.js'
import { areaText, parseRange } from '../src/structure/area.js'
import { seeded } from './looks-book.js'

const [ROWS, COLS] = [14, 12]

/** The look of each cell of the corner, painted rectangle by rectangle (`style|format`, empty for none). */
function painted(list: LookRange[]): Map<string, string> {
  const out = new Map<string, string>()
  for (const x of list) {
    const r = parseRange(x.range)!
    for (let row = r.r1; row <= Math.min(r.r2, ROWS + 2); row++)
      for (let c = r.c1; c <= Math.min(r.c2, COLS + 2); c++) out.set(`${row}:${c}`, x.style === undefined && x.numberFormat === undefined ? '' : `${x.style ?? ''}|${x.numberFormat ?? ''}`)
  }
  return out
}

/** The layer's look of each cell of the corner, as `painted` writes it. */
function shown(layer: Layer): Map<string, string> {
  const out = new Map<string, string>()
  for (let r = 0; r <= ROWS + 2; r++)
    for (let c = 0; c <= COLS + 2; c++) {
      const l = layer.lookAt(r, c)
      out.set(`${r}:${c}`, l ? `${l.style ?? ''}|${l.numberFormat ?? ''}` : '')
    }
  return out
}

const without = (m: Map<string, string>) => new Map([...m].filter(([, v]) => v !== ''))

/** Whether two layers give every cell the same look. */
function sameLooks(a: Layer, b: Layer): boolean {
  const [x, y] = [[...a.pieces()], [...b.pieces()]]
  if (x.length !== y.length) return false
  return x.every((p, i) => {
    const q = y[i]
    return p.r1 === q.r1 && p.r2 === q.r2 && p.c1 === q.c1 && p.c2 === q.c2 && sameLook(a.looks.list[p.look], b.looks.list[q.look])
  })
}

/** A random list of rectangles: cells, ranges, whole rows and columns, of two styles and two formats, or of none. */
function randomList(rnd: (n: number) => number): LookRange[] {
  return Array.from({ length: rnd(9) }, () => {
    const [r1, c1] = [rnd(ROWS), rnd(COLS)]
    const kind = rnd(10)
    const x = kind === 0 ? { r1, r2: r1 + rnd(3), c1: 0, c2: MAX_COLS - 1 } : kind === 1 ? { r1: 0, r2: MAX_ROWS - 1, c1, c2: c1 + rnd(2) } : { r1, r2: r1 + rnd(5), c1, c2: c1 + rnd(6) }
    const look = rnd(4)
    return { range: areaText(x), ...(look & 1 ? { style: rnd(2) } : {}), ...(look & 2 ? { numberFormat: rnd(2) ? '0.0' : 14 } : {}) }
  })
}

describe('the layer of a sheet\'s looks', () => {
  it('a list read gives each cell the look of the last rectangle over it; written in its form, it reads back the same', () => {
    const rnd = seeded(11)
    for (let k = 0; k < 2000; k++) {
      const list = randomList(rnd)
      const layer = readRanges(list)
      expect(without(shown(layer)), JSON.stringify(list)).toEqual(without(painted(list)))
      const again = readRanges(writeRanges(layer))
      expect(sameLooks(again, layer), JSON.stringify(list)).toBe(true)
      // the form is one for one set of looks: written again from what it reads, the same list
      expect(writeRanges(again)).toEqual(writeRanges(layer))
    }
  })

  it('cells written over a layer take their looks, a cell of none clears its look', () => {
    const rnd = seeded(5)
    for (let k = 0; k < 1000; k++) {
      const list = randomList(rnd)
      const layer = readRanges(list)
      const want = painted(list)
      const cells: [number, number, number][] = Array.from({ length: rnd(8) }, () => {
        const [r, c, look] = [rnd(ROWS), rnd(COLS), rnd(3)]
        const id = look === 0 ? NONE : layer.looks.id(look - 1, undefined)
        want.set(`${r}:${c}`, id === NONE ? '' : `${look - 1}|`)
        return [r, c, id]
      })
      const next = edited(layer, cellEdits(cells))
      expect(without(shown(next))).toEqual(without(want))
      expect(sameLooks(readRanges(writeRanges(next)), next)).toBe(true)
    }
  })

  it('rows of one pattern apart from each other, with rows between them that differ a little, are written over one base', () => {
    // rows in threes with two empty rows between, each of style 0 with style 1 in every 12th column, as the wide shape's sheet
    const list: LookRange[] = []
    for (let r = 20; r < 200; r += 5) {
      list.push({ range: `${r}:${r + 2}`, style: 0 })
      for (let c = 17; c < 600; c += 12) list.push({ range: `${formatAddress(r - 1, c)}:${formatAddress(r + 1, c)}`, style: 1, numberFormat: '0.0%' })
    }
    // a cell of none in one row (as under an entry with content), and a cell of another look in another
    list.push({ range: 'C50' }, { range: 'D60', numberFormat: '0' })
    const layer = readRanges(list)
    expect(layer.blocks).toBe(36 + 2)
    const out = writeRanges(layer)
    // the base's runs across the rows of the group, the rows between as rows of none, the two rows' differences
    const runs = 2 * Math.ceil((600 - 17) / 12) + 1
    expect(out.slice(0, runs).every((x) => /^[A-Z]+20:[A-Z]+197$/.test(x.range))).toBe(true)
    expect(out.length).toBe(runs + 35 + 2)
    expect(sameLooks(readRanges(out), layer)).toBe(true)
  })

  it('lines inserted and deleted, rectangles cut and pasted move the layer as they move its cells; the cells that differ are found', () => {
    const rnd = seeded(23)
    const key = (r: number, c: number) => `${r}:${c}`
    for (let k = 0; k < 1500; k++) {
      const layer = readRanges(randomList(rnd).filter((x) => !/^[A-Z]+:[A-Z]+$/.test(x.range)))
      const was = without(shown(layer))
      const look = (r: number, c: number) => {
        const l = layer.lookAt(r, c)
        return l ? `${l.style ?? ''}|${l.numberFormat ?? ''}` : undefined
      }
      const want = new Map<string, string>()
      const set = (r: number, c: number, v: string | undefined) => v !== undefined && r >= 0 && c >= 0 && want.set(key(r, c), v)
      let got: Layer | undefined
      const [axis, at, n] = [rnd(2) ? 'r' : 'c', rnd(ROWS), rnd(2) ? 1 + rnd(3) : -1 - rnd(3)] as const
      const move = rnd(3)
      if (move < 2) {
        // lines: a cell before `at` stays, one after moves by n, one deleted is gone; the new lines copy the line before them
        const before = move === 1 && n > 0
        got = shifted(layer, axis, at, n, before)
        // past the corner too: deleted lines bring cells from there into it
        for (let r = 0; r <= ROWS + 6; r++)
          for (let c = 0; c <= COLS + 6; c++) {
            const i = axis === 'r' ? r : c
            const to = i < at ? i : n > 0 ? i + n : i >= at - n ? i + n : -1
            if (to >= 0) set(axis === 'r' ? to : r, axis === 'r' ? c : to, look(r, c))
            if (before && i === at - 1) for (let j = at; j < at + n; j++) set(axis === 'r' ? j : r, axis === 'r' ? c : j, look(r, c))
          }
      } else {
        // a range cut and pasted on the same sheet: the source emptied, the target the source's
        const src = { r1: rnd(ROWS), c1: rnd(COLS), r2: 0, c2: 0 }
        Object.assign(src, { r2: src.r1 + rnd(4), c2: src.c1 + rnd(4) })
        const [dr, dc] = [rnd(9) - 4, rnd(9) - 4]
        got = carried([layer], [{ kind: 'range', from: 0, src, to: 0, dr, dc }])[0]
        for (const [p, v] of was) want.set(p, v)
        for (let r = src.r1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) want.delete(key(r, c))
        for (let r = src.r1 + dr; r <= src.r2 + dr; r++) for (let c = src.c1 + dc; c <= src.c2 + dc; c++) want.delete(key(r, c))
        for (let r = src.r1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) set(r + dr, c + dc, look(r, c))
      }
      expect(without(shown(got!))).toEqual(new Map([...want].filter(([p]) => Number(p.split(':')[0]) <= ROWS + 2 && Number(p.split(':')[1]) <= COLS + 2)))
      // the cells whose look differs from the layer before, found by blocks and runs
      const differs: string[] = []
      for (const x of differing(layer, got)) for (let r = x.r1; r <= Math.min(x.r2, ROWS + 2); r++) for (let c = x.c1; c <= Math.min(x.c2, COLS + 2); c++) differs.push(key(r, c))
      const now = without(shown(got!))
      expect(new Set(differs)).toEqual(new Set([...new Set([...was.keys(), ...now.keys()])].filter((p) => was.get(p) !== now.get(p))))
    }
  })

  it('a sheet is taken with its entries of a look alone over its rectangles, its entries of content holding none under them', () => {
    const sheet = looksTaken({
      name: 'S',
      cells: { A1: { value: 1 }, B1: { style: 0 }, B2: { style: 1, numberFormat: '0.0' }, C1: {}, C2: { note: { text: 'n' } } },
      looks: [{ range: 'A1:C3', style: 2 }, { range: 'C3' }],
    })
    expect(sheet.cells).toEqual({ A1: { value: 1 }, C1: {}, C2: { note: { text: 'n' } } })
    const layer = readRanges(sheet.looks!)
    const look = (a: string) => {
      const x = parseRange(a)!
      return layer.lookAt(x.r1, x.c1)
    }
    expect(['A1', 'B1', 'B2', 'C1', 'C2', 'A3', 'C3', 'D1'].map(look)).toEqual([undefined, { style: 0 }, { style: 1, numberFormat: '0.0' }, undefined, undefined, { style: 2 }, undefined, undefined])
    // taken again, it is the same sheet
    expect(looksTaken(sheet)).toBe(sheet)
  })
})
