import { describe, expect, it } from 'vitest'
import { createSession } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { partsOf, withParts } from './keep-files.js'

const none = '<fill><patternFill patternType="none"/></fill>'
const gray = '<fill><patternFill patternType="gray125"/></fill>'
const cases = [
  { name: 'missing fills', xml: '', entries: [] },
  { name: 'empty fills', xml: '<fills count="0"/>', entries: [] },
  { name: 'one reserved fill', xml: `<fills count="1">${none}</fills>`, entries: [none] },
  { name: 'both reserved fills', xml: `<fills count="2">${none}${gray}</fills>`, entries: [none, gray] },
]

describe('writeKept: Excel reserves fill indices 0 and 1 (MS-OI29500 2.1.705)', () => {
  it.each(cases)('adds a visible fill after $name, preserving existing entries', async ({ xml, entries }) => {
    const data = await withParts({ sheets: [{ name: 'S', cells: { A1: { value: 5 }, B1: { value: 6 }, C1: { value: 7 } } }] }, {
      edit: { 'xl/styles.xml': (styles) => styles.replace(/<fills\b[^>]*>[\s\S]*?<\/fills>/, xml) },
    })
    const opened = await openKept(data)
    const session = createSession(opened.book, { cached: opened.cached })
    // Merely opening an incomplete fills list must not rewrite the file.
    expect((await writeKept(data, session.getSnapshot())).data).toEqual(data)
    session.patchCells([
      { sheet: 'S', range: 'A1:B1', set: { style: { fill: { color: '#FF0000' } } } },
      { sheet: 'S', cell: 'C1', set: { style: { fill: { color: '#00FF00' } } } },
    ])
    const written = await writeKept(data, session.getSnapshot())
    const { book } = await readXlsx(written.data)
    for (const [ref, color] of [['A1', '#FF0000'], ['B1', '#FF0000'], ['C1', '#00FF00']]) {
      const cell = book.sheets[0].cells[ref]
      expect(book.styles?.[cell.style!]?.fill?.color, ref).toBe(color)
    }
    const styles = (await partsOf(written.data)).get('xl/styles.xml')!
    const fills = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(styles)![1].match(/<fill>.*?<\/fill>/g)!
    expect(fills.slice(0, entries.length)).toEqual(entries)
    expect(fills.slice(0, 2)).toEqual([none, gray])
    expect(fills).toHaveLength(4) // Reserved entries, one red fill shared by A1/B1, and one green.
    expect(written.issues).toEqual([])
  })
})
