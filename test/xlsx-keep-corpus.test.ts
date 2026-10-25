/**
 * `xlcalc/xlsx/keep` over the files of the test sets (every third of them, `test/excel`): opened and written without an
 * edit, a file is its own bytes; after random edits (values, text, formulas, styles, merged areas, rules, validation,
 * hyperlinks, and operations of the structure), read back it is the book `writeXlsx` writes read back, and the parts
 * no edit concerns (the other sheets, the parts the model does not have) are the file's byte for byte. A change the file
 * cannot take is refused (`KeptRefusal`), never written.
 */
import { execSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { createSession, registerLocale, type BookInput } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { keyOf } from '../src/style-table.js'
import { parse } from '../src/syntax/parse.js'
import { canonical } from '../src/syntax/print.js'
import { readXlsx, writeXlsx, type XlsxRead } from '../src/xlsx.js'
import { BUILTIN } from '../src/xlsx/styles/numfmt.js'
import { relsOf } from '../src/xlsx/read/package.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'

registerLocale(ruRU)
const root = fileURLToPath(new URL('..', import.meta.url))
const files = execSync("find test/excel -name '*.xlsx'", { cwd: root, encoding: 'utf8' }).split('\n').filter(Boolean).sort().filter((_, i) => i % 3 === 0)

/**
 * Where the file written over the original reads back otherwise than `writeXlsx`'s, as the original has it: a view's
 * selection the writer fills in, collapsed marks the writer writes as Excel keeps them, a formula the writer cannot write.
 */
const KNOWN: Record<string, string> = {
  'made-frozen.xlsx': '.view',
  'hand-3.xlsx': '.rows',
  'let-1.xlsx': '.formula',
  'let-held-1.xlsx': '.formula',
}

/** A formula canonical, so that `L!A1` (as Excel writes it) and `'L'!A1` (as the writer does) agree. */
const canon = (f: string) => {
  try {
    return canonical(parse(f, { union: true }))
  } catch {
    return f
  }
}
const FORMULAS = new Set(['formula', 'formula1', 'formula2', 'ref'])
function formulas(v: unknown): unknown {
  if (Array.isArray(v)) return v.map(formulas)
  if (v === null || typeof v !== 'object') return v
  return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, typeof x === 'string' && FORMULAS.has(k) ? canon(x) : k === 'formulas' && Array.isArray(x) ? x.map(canon) : formulas(x)]))
}

/**
 * A read book with its styles by their fields, its formulas canonical, a rule's built-in number format by its code, and
 * its active sheet by name (a file with chart sheets names its first worksheet, which a file without leaves out).
 */
function norm(r: XlsxRead) {
  const st = r.book.styles ?? []
  const key = (i: number | undefined) => (i === undefined ? undefined : keyOf(st[i] ?? {}))
  const line = <T extends { style?: number }>(e: T) => (e.style !== undefined ? { ...e, style: key(e.style) } : e)
  const map = <T extends { style?: number }>(o: Record<string, T> | undefined) => o && Object.fromEntries(Object.entries(o).map(([k, v]) => [k, line(v)]))
  const { styles: _, ...book } = r.book
  return {
    book: formulas({
      ...book,
      activeSheet: book.activeSheet ?? book.sheets.find((x) => !x.hidden)?.name,
      sheets: book.sheets.map((s) => ({
        ...s,
        cells: map(s.cells),
        rows: map(s.rows),
        columns: map(s.columns),
        looks: s.looks?.map(line),
        conditionalFormats: s.conditionalFormats?.map((c) => ('style' in c && typeof c.style?.numberFormat === 'number' ? { ...c, style: { ...c.style, numberFormat: BUILTIN[c.style.numberFormat] ?? c.style.numberFormat } } : c)),
      })),
    }),
    cached: r.cached,
  }
}

/** The first place two values differ, as a path. */
function firstDiff(a: unknown, b: unknown, path = ''): string | undefined {
  if (a === b) return undefined
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return `${path}: ${JSON.stringify(a)?.slice(0, 120)} vs ${JSON.stringify(b)?.slice(0, 120)}`
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
    const [x, y] = [(a as Record<string, unknown>)[k], (b as Record<string, unknown>)[k]]
    if (x === undefined && y === undefined) continue
    const d = firstDiff(x, y, `${path}.${k}`)
    if (d) return d
  }
  return undefined
}

const col = (i: number) => {
  let s = ''
  for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s
  return s
}

const EDITS = ['value', 'text', 'formula', 'style', 'merge', 'cond', 'valid', 'link'] as const
const STRUCTURE = ['value', 'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns', 'moveRange', 'renameSheet', 'insertSheet', 'deleteSheet', 'moveSheet'] as const

/** One to four random edits of a kind, seeded by the file's path; the sheets of the book read the edits concern. */
function edit(file: string, read: XlsxRead, kinds: readonly string[]) {
  let seed = [...file].reduce((h, c) => (h * 31 + c.charCodeAt(0)) & 0x7fffffff, 13)
  const rnd = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff
  const pick = <T,>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)]
  const s = createSession(read.book, { cached: read.cached })
  const sheets = read.book.sheets.filter((x) => !x.hidden)
  for (let e = 0, n = 1 + Math.floor(rnd() * 4); e < n && sheets.length; e++) {
    const sh = pick(sheets)
    const used = s.getUsedRange(sh.name) as { rows?: number; columns?: number } | null
    const [rows, cols] = [Math.max(1, used?.rows ?? 10) + 2, Math.max(1, used?.columns ?? 5) + 2]
    const at = () => `${col(Math.floor(rnd() * Math.min(cols, 30)))}${1 + Math.floor(rnd() * Math.min(rows, 200))}`
    const kind = pick(kinds)
    const cell = at()
    const now = () => s.getBook().sheets.find((x) => x.name === sh.name)
    try {
      const r = (() => {
        switch (kind) {
          case 'value': return s.setCells([{ sheet: sh.name, cell, input: { value: Math.round(rnd() * 1000) / 10 } }])
          case 'text': return s.setCells([{ sheet: sh.name, cell, input: { value: `kept ${e}` } }])
          case 'formula': return s.setCells([{ sheet: sh.name, cell, input: { formula: `SUM(${at()}:${at()})*2` } }])
          case 'style': return s.patchCells([{ sheet: sh.name, cell, set: { style: pick([{ font: { bold: true } }, { fill: { color: '#FFEE00' } }, { border: { bottom: { style: 'thin' } } }]) } }])
          case 'merge': return s.editStructure({ op: 'mergeCells', sheet: sh.name, range: `${cell}:${col(Math.min(cols + 3, 40))}${Number(cell.replace(/\D/g, '')) + 1}` })
          case 'cond': return s.setLayout([{ sheet: sh.name, conditionalFormats: [...(now()?.conditionalFormats ?? []), { type: 'cellIs', ranges: [cell], operator: 'greaterThan', formulas: ['5'], style: { font: { bold: true }, fill: { color: '#FF0000' } } }] }])
          case 'valid': return s.setLayout([{ sheet: sh.name, validations: [...(now()?.validations ?? []).filter((v) => !v.range.split(' ').includes(cell)), { range: cell, type: 'whole', operator: 'greaterThan', formula1: '0' }] }])
          case 'link': return s.patchCells([{ sheet: sh.name, cell, set: { link: { url: 'https://example.com/kept' } } }])
          case 'insertRows':
          case 'deleteRows': return s.editStructure({ op: kind, sheet: sh.name, row: 1 + Math.floor(rnd() * Math.min(rows, 50)), count: 1 + Math.floor(rnd() * 3) })
          case 'insertColumns':
          case 'deleteColumns': return s.editStructure({ op: kind, sheet: sh.name, column: col(Math.floor(rnd() * Math.min(cols, 10))), count: 1 })
          case 'moveRange': return s.editStructure({ op: 'moveRange', from: { sheet: sh.name, range: `${cell}:${at()}` }, to: { sheet: sh.name, cell: at() } })
          case 'renameSheet': return s.editStructure({ op: 'renameSheet', sheet: sh.name, name: `${sh.name.slice(0, 20)} r${e}` })
          case 'insertSheet': return s.editStructure({ op: 'insertSheet', name: `New ${e}`, before: sh.name })
          case 'deleteSheet': return s.editStructure({ op: 'deleteSheet', sheet: sh.name })
          default: return s.editStructure({ op: 'moveSheet', sheet: sh.name })
        }
      })()
      if ((r as { ok?: boolean }).ok !== false && /Sheet$/.test(kind)) break
    } catch {
      // an edit the session refuses
    }
  }
  const snap = s.getSnapshot()
  s.close()
  return snap
}

/** The sheets of the book read whose input or values the snapshot changed, the active one's tab among them: by their carriers. */
function concerned(read: XlsxRead, snap: { book: BookInput; result: { sheets: { cells: Record<string, { value: unknown }> }[] } }): Set<number> {
  const out = new Set<number>()
  const activeOf = (b: BookInput) => b.activeSheet ?? b.sheets.find((x) => !x.hidden)?.name
  const index = new Map(snap.book.sheets.map((x, j) => [x.name.toUpperCase(), j]))
  const strip = ({ name: _, ...rest }: BookInput['sheets'][number]) => keyOf(rest)
  read.book.sheets.forEach((sh, i) => {
    const carrier = snap.book.names?.find((n) => n.name === `_xlcalc.keep.s${i}`)
    const m = carrier && /^'((?:[^']|'')+)'!/.exec(carrier.ref)
    const j = m ? index.get(m[1].replace(/''/g, "'").toUpperCase()) : undefined
    const now = j === undefined ? undefined : snap.result.sheets[j].cells
    const was = read.cached.sheets[i].cells
    if (j === undefined || strip(snap.book.sheets[j]) !== strip(sh) || keyOf(snap.book.styles ?? []) !== keyOf(read.book.styles ?? [])) out.add(i)
    else if (activeOf(read.book) !== activeOf(snap.book) && [activeOf(read.book), activeOf(snap.book)].includes(sh.name)) out.add(i)
    else if (Object.keys(now!).some((a) => keyOf(now![a].value) !== keyOf(was[a]?.value)) || Object.keys(was).some((a) => !now![a])) out.add(i)
  })
  return out
}

/** Parts with references the model does not hold, which a structure operation moves, and their text without those references. */
const MOVED = /^xl\/(charts\/chart|drawings\/drawing|pivotTables\/|pivotCache\/pivotCacheDefinition)/i
const refsOut = (xml: string) => xml.replace(/<(\w+:)?f(\s[^>]*)?>[^<]*<\/(\w+:)?f>/g, '<f/>').replace(/<xdr:(col|row|colOff|rowOff)>\d+</g, '<xdr:$1><').replace(/\s(ref|sheet)="[^"]*"/g, '')

const PARTS_OF_BOOK = /^(\[content_types\]\.xml|xl\/workbook\.xml|xl\/_rels\/workbook\.xml\.rels|xl\/styles\.xml|xl\/calcchain\.xml|xl\/metadata\.xml|xl\/richdata\/.*)$/i

/** Each file edited by `kinds`, written over its original; what differs from writeXlsx's, and the parts that changed unasked. */
async function check(kinds: readonly string[]): Promise<{ written: number; refused: number; problems: string[] }> {
  const out = { written: 0, refused: 0, problems: [] as string[] }
  for (const f of files) {
    const data = readFileSync(`${root}/${f}`)
    const read = await openKept(data)
    const snap = edit(f, read, kinds)
    let w
    try {
      w = await writeKept(data, snap)
    } catch (e) {
      if ((e as Error).name !== 'KeptRefusal') throw new Error(`${f}: ${(e as Error).stack}`)
      out.refused++
      continue
    }
    out.written++
    if (Buffer.compare(Buffer.from(w.data), data) === 0) continue
    const plain = { ...snap, book: { ...snap.book, names: snap.book.names?.filter((n) => !n.name.startsWith('_xlcalc.keep.')) } }
    if (!plain.book.names?.length) delete plain.book.names
    const d = firstDiff(norm(await readXlsx(w.data)), norm(await readXlsx((await writeXlsx(plain)).data)))
    const known = KNOWN[f.slice(f.lastIndexOf('/') + 1)]
    if (d && !(known && d.includes(known))) out.problems.push(`${f}: ${d}`)
    // the parts of the sheets no edit concerns, and the parts the model does not have, are the file's
    const [o, x] = [openArchive(data), openArchive(w.data)]
    const { openSource } = await import('../src/xlsx/keep/source.js')
    const src = await openSource(data)
    // a concerned sheet's part and the parts under it (a sheet deleted takes its drawings with their charts)
    const mine = new Set<string>()
    const under = async (part: string) => {
      if (mine.has(part.toLowerCase())) return
      mine.add(part.toLowerCase()).add(part.replace(/([^/]+)$/, '_rels/$1.rels').toLowerCase())
      for (const r of await relsOf(src.archive, part)) if (!r.external) await under(r.target)
    }
    for (const i of concerned(read, snap)) await under(src.sheets[i].part)
    for (const name of o.names) {
      if (PARTS_OF_BOOK.test(name) || mine.has(name.toLowerCase())) continue
      const [p, q] = [await o.bytes(name), await x.bytes(name)]
      if (q && Buffer.compare(Buffer.from(p!), Buffer.from(q)) === 0) continue
      // a part whose references an operation moved: it changed in them alone
      if (q && MOVED.test(name) && refsOut((await o.text(name))!) === refsOut((await x.text(name))!)) continue
      out.problems.push(`${f}: ${name} ${q ? 'changed' : 'missing'}`)
    }
  }
  return out
}

describe('xlsx/keep over the test sets', () => {
  it('a file opened and written without an edit is its own bytes', async () => {
    const bad: string[] = []
    for (const f of files) {
      const data = readFileSync(`${root}/${f}`)
      const r = await openKept(data)
      const w = await writeKept(data, createSession(r.book, { cached: r.cached }).getSnapshot())
      if (Buffer.compare(Buffer.from(w.data), data) !== 0) bad.push(f)
    }
    expect(files.length).toBeGreaterThan(300)
    expect(bad).toEqual([])
  }, 120_000)

  it('random edits: read back as writeXlsx writes the book, the other parts as they were', async () => {
    const r = await check(EDITS)
    expect(r.problems).toEqual([])
    expect(r.refused).toBe(0)
  }, 240_000)

  it('random operations of the structure: written as writeXlsx writes the book, or refused', async () => {
    const r = await check(STRUCTURE)
    expect(r.problems).toEqual([])
    expect(r.written).toBeGreaterThan(r.refused * 20)
  }, 240_000)
})
