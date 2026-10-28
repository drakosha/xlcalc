/**
 * Session reads that make their parts once for many cells: the looks of a large window's listed cells, the cells of a
 * page of readCells and findCells, the values of summarizeRanges, the source cell of many errors over one range, the
 * rules of data validation over many typed cells. Each gives what the same read of one cell at a time gives.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput } from '../src/contract.js'
import { createSession } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const RED = { fill: { color: '#FF0000' } }

describe.each(['closures', 'eval'] as const)('reads together, %s', (execution) => {
  it("a large window's listed cells take their rules' looks as each cell alone does, a rule that stops the rest too", () => {
    const cells: BookInput['sheets'][number]['cells'] = {}
    for (let k = 0; k < 300; k++) cells[`${String.fromCharCode(65 + (k % 26))}${k * 331 + 1}`] = { value: k % 40 }
    const s = createSession({ sheets: [{ name: 'S', cells, conditionalFormats: [
      { type: 'cellIs', ranges: ['A:C', 'E:Z'], operator: 'greaterThan', formulas: ['30'], style: RED, stopIfTrue: true },
      { type: 'expression', ranges: ['A1:Z100000'], formula: 'MOD(A1,2)=0', style: { font: { bold: true }, numberFormat: '0.00' } },
    ] }] }, { clock, execution, display: true })
    const window = s.getCells('S', 'A1:Z100000')
    expect(Object.keys(window)).toHaveLength(300)
    for (const a of Object.keys(cells)) expect(window[a]).toEqual(s.getCell('S', a))
    expect(Object.values(window).filter((v) => v.appearance?.style?.fill).length).toBeGreaterThan(0)
    const page = s.readCells([{ sheet: 'S', range: 'A:Z' }], { maxCells: 1000, fields: ['text'] })
    for (const c of page.cells) expect(c.text).toBe(s.getCell('S', c.cell)!.text)
    const found = s.findCells([{ sheet: 'S', range: 'A:Z' }], '.00', { maxCells: 1000, fields: ['text'] })
    expect(found.cells.length).toBeGreaterThan(0)
    for (const c of found.cells) expect(c.text).toBe(s.getCell('S', c.cell)!.text)
  })

  it('summarizeRanges counts each cell once, by the rows asked, past a cell that is gone and with spilled cells', () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 'x' }, A4: { formula: '1/0' }, B1: { value: 10 }, D1: { formula: 'SEQUENCE(3)' } } }] }, { clock, execution, display: false })
    const sum = (ranges: string[], rows?: number[]) => s.summarizeRanges(ranges.map((range) => ({ sheet: 'S', range })), rows ? { rows } : undefined)
    expect(sum(['A1:B4', 'A1:A2'])).toMatchObject({ cells: 5, numbers: 3, errors: 1, sum: 13, min: 1, max: 10 })
    expect(sum(['A1:B4'], [2, 3])).toMatchObject({ cells: 2, numbers: 1, sum: 2 })
    expect(sum(['A1:D4'])).toMatchObject({ cells: 8, numbers: 6, sum: 19 })
    s.setCells([{ sheet: 'S', cell: 'A2', input: null }])
    expect(sum(['A1:B4'])).toMatchObject({ cells: 4, numbers: 2, sum: 11 })
  })

  it('formulas with an error over one range name the first cell of it that holds their error', () => {
    const cells: BookInput['sheets'][number]['cells'] = { A3: { formula: 'NA()' }, A7: { formula: '1/0' }, A9: { formula: 'NA()' } }
    for (let r = 1; r <= 6; r++) cells[`C${r}`] = { formula: r % 2 ? 'SUM($A$1:$A$10)' : 'MATCH(1/0,$A$1:$A$10,0)+SUM($A$1:$A$10)' }
    const s = createSession({ sheets: [{ name: 'S', cells }] }, { clock, execution, display: false })
    const reasons = s.getResult().diagnostics.flatMap((i) => (i.kind === 'cell' && i.cell.startsWith('C') ? [i.reason] : []))
    // each error by its own code: #N/A from A3, #DIV/0! from A7, which a range holds after an #N/A
    const na = "value not found, from 'S'!A3"
    const div = "division by zero, from 'S'!A7"
    expect(reasons).toEqual([na, div, na, div, na, div])
    expect(s.getCellIssue('S', 'C2')!.reason).toBe(div)
  })

  it("typed texts meet the rule of their cell's range, after a structure operation moves the range too", () => {
    const validations = Array.from({ length: 30 }, (_, i) => ({ range: `B${i + 1}`, type: 'whole' as const, operator: 'greaterThan' as const, formula1: '10' }))
    const s = createSession({ sheets: [{ name: 'S', cells: {}, validations: [...validations, { range: 'D1:D5 F2', type: 'list', formula1: '"a,b"' }] }] }, { clock, execution, display: false })
    const alerts = (cells: string[], text: string) => {
      const p = s.previewCellTexts(cells.map((cell) => ({ sheet: 'S', cell, text })))
      return p.ok ? (p.validations ?? []).map((v) => v.cell) : p
    }
    expect(alerts(['B3', 'B31', 'D4', 'F2', 'F3'], '5')).toEqual(['B3', 'D4', 'F2'])
    s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(alerts(['B1', 'B3', 'B31', 'D1', 'F3'], '5')).toEqual(['B3', 'B31', 'F3'])
  })
})
