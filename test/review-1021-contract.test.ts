/**
 * The contract's types name what the calls return: the refusals of typed text include a reference to a workbook not among
 * the links. The views of a cell give a host copies of its note and link in every kind of session.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { createSession, type BookInput } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const OTHER = '=[Book9.xlsx]Sheet1!A1'

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  it("the calls that type text refuse a workbook not among the links with 'external_ref', a code of their result types", () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'x' } } }] }, { clock: { now: '2026-10-04T12:00:00Z', seed: 7 }, execution })
    const set = s.setCellText('S', 'B1', OTHER)
    const preview = s.previewCellText('S', 'B1', OTHER)
    const texts = s.setCellTexts([{ sheet: 'S', cell: 'C1', text: OTHER }])
    const textsPreview = s.previewCellTexts([{ sheet: 'S', cell: 'C1', text: OTHER }])
    const fill = s.fillCellText('S', 'D1:D2', OTHER)
    const fillPreview = s.previewFillCellText('S', 'D1:D2', OTHER)
    const replaced = s.replaceInCells([{ sheet: 'S', range: 'A1' }], 'x', OTHER)
    // each comparison compiles only where the result's code can be 'external_ref'
    expect([
      !set.ok && set.code === 'external_ref',
      !preview.ok && preview.code === 'external_ref',
      !texts.ok && texts.code === 'external_ref',
      !textsPreview.ok && textsPreview.code === 'external_ref',
      !fill.ok && fill.code === 'external_ref',
      !fillPreview.ok && fillPreview.code === 'external_ref',
      replaced.refused?.[0].code === 'external_ref',
    ]).toEqual(Array(7).fill(true))
    s.close()
  })
})

describe.each(['closures', 'eval'] as const)('views of a cell hand out copies of its note and link (fix-u319), %s', (execution) => {
  afterAll(closeCalcs)
  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('changing them leaves the workbook as it was', async () => {
      const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1, note: { text: 'n' }, link: { url: 'http://a' } } } }] }
      const s = await open(book, { execution, clock: { now: '2026-10-04T12:00:00Z', seed: 7 } })
      const hash = (await s.getResult()).inputHash
      const cell = (await s.getCell('S', 'A1'))!
      const cells = await s.getCells('S', 'A1')
      const vp = await s.getViewport('S', 'A1:B2')
      cell.note!.text = 'edited'
      cells.A1.link!.url = 'http://cells'
      vp.cells.A1.note!.text = 'viewport'
      vp.cells.A1.link!.url = 'http://viewport'
      const now = (await s.getBook()).sheets[0].cells.A1
      expect([now.note, now.link]).toEqual([{ text: 'n' }, { url: 'http://a' }])
      expect((await s.getResult()).inputHash).toBe(hash)
      await s.close()
    })
  })
})
