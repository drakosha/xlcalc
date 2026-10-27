/**
 * Paste Special (`copyRange` with `paste` and `transpose`), as Excel for Mac pastes it (the paste set `probe`). A kind
 * of paste lays a part of the copy over the target and leaves the rest of the target as it was: the copy is made whole,
 * then each cell of the target takes its contents from the copy or the target, and its look likewise. A transposed copy
 * is made from a scratch sheet holding the source turned rows into columns where the target is: Excel moves every
 * reference of the source's formulas as the cell it points to turns about the source's top-left cell, an absolute one
 * too, and swaps its `$` between the axes (`B2`'s `$A$1+A$2` copied from A1 to F1 is `$F$1+$G1`).
 */
import type { BookInput, CellInput, ConditionalFormat, DataValidation, PasteKind, SheetInput } from '../contract.js'
import { mapFormulas } from '../cond/structure.js'
import { originOf, rectsOf } from '../cond/rules.js'
import { spend } from '../exec/budget.js'
import { lookOf } from '../layout/lines.js'
import { mergesOf, parseMerge } from '../layout/merges.js'
import { entryAt, lines } from '../layout/spans.js'
import { say } from '../locale/messages.js'
import { cellEdits, edited, LookTable, NONE } from '../looks/layer.js'
import { expanded, layerOf, lookIn, withField, withLayer } from '../looks/sheet.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { sheetKey } from '../syntax.js'
import { anchorOf, areasOf, rangeText } from '../validation/rules.js'
import { withFormulas } from '../validation/structure.js'
import { cellCount, contains, cornerOf, cornersText, inside, isFull, LIMIT, meets, parseRange, rectText, shiftRect, span, type Area, type Axis, type Rect, type Span } from './area.js'
import { clipboard, pasteArea, pasteClip, type Clipboard } from './filtered.js'
import { copyTables, headersPasted } from '../tables/moves.js'
import type { ValuesIn } from './values.js'
import { chain, type Applied } from './geometry.js'
import { tiles } from './copy.js'
import { cellsIn, gridOf } from './grid.js'
import { rewriteFormula, sheetIndex, sheetNames, sheetOf, type Transform } from './rewrite.js'
import { formulaAt } from './transform.js'

type Refused = Exclude<Applied, { ok: true }>

/** The scratch sheet of a transposed copy, which no sheet of a book can have. */
export const TURNED = '\u0000turned'

/** Whether a kind of paste pastes the cells' contents as values (from a clipboard of values, `clipped`). */
export const ofValues = (kind: PasteKind) => kind === 'values'

/** What a kind of paste takes from the copy; the rest stays the target's. */
interface Parts {
  /** Values or formulas, and `meta` with them. */
  content: boolean
  /** Styles, number formats, the lines of a whole-line copy, merged areas and conditional formats. */
  looks: boolean
  validation: boolean
}

const PARTS: Record<Exclude<PasteKind, 'all'>, Parts> = {
  values: { content: true, looks: false, validation: false },
  formulas: { content: true, looks: false, validation: false },
  formats: { content: false, looks: true, validation: false },
  validation: { content: false, looks: false, validation: true },
}

/** A cell's fields that hold what it contains: Paste Special takes them or leaves them together. */
const CONTENT = ['value', 'formula', 'array', 'held', 'meta'] as const

/** Whether the paste writes the cells' contents (and so is refused onto merged areas and renames header cells). */
export const writesContent = (kind: PasteKind) => kind === 'all' || PARTS[kind].content

/**
 * One axis of a reference turned onto the other: from the source's `from` to the target's `to` on the other axis, its
 * `$` kept; a whole line's span starts where its first line turns and runs to the other axis's end (Excel's `A:A` from
 * A1 to F1 is `$F1:$XFD1`, `5:5` from D2 to G4 is `J:J`). Null when it turns off the sheet.
 */
function turnedSpan(s: Span, axis: Axis, from: number, to: number): Span | null {
  const other: Axis = axis === 'r' ? 'c' : 'r'
  if (isFull(s, axis)) return { lo: { i: Math.max(0, to - from), abs: s.lo.abs }, hi: { i: LIMIT[other] - 1, abs: s.hi.abs } }
  const out = span({ i: to + s.lo.i - from, abs: s.lo.abs }, { i: to + s.hi.i - from, abs: s.hi.abs })
  return out.lo.i < 0 || out.hi.i >= LIMIT[other] ? null : out
}

/** Whether a reference has a `$` of its own: a whole line's span has none written. */
const anchored = (area: Area) => (['r', 'c'] as const).some((axis) => !isFull(area[axis], axis) && (area[axis].lo.abs || area[axis].hi.abs))

/**
 * References turned about `src`'s top-left cell on sheet `from` onto `at`, the target's (the paste set `held`): a
 * reference into the source, or one without a `$` anywhere, goes `dc` rows and `dr` columns from the target's corner
 * where it was `dr` rows and `dc` columns from the source's, its `$` swapping axes with it; any other reference with a
 * `$` stays. A reference turned off the sheet is lost.
 */
function turnTransform(from: string, src: Rect, at: { r: number; c: number }): Transform {
  const key = sheetKey(from)
  return {
    copy: true,
    ref({ sheet, area }) {
      const own = sheetKey(sheet) === key && contains(src, { r1: area.r.lo.i, r2: area.r.hi.i, c1: area.c.lo.i, c2: area.c.hi.i })
      if (!own && anchored(area)) return undefined
      const [r, c] = [turnedSpan(area.c, 'c', src.c1, at.r), turnedSpan(area.r, 'r', src.r1, at.c)]
      if (!r || !c) return 'lost'
      const kind = isFull(r, 'r') ? 'cols' : isFull(c, 'c') ? 'rows' : area.kind === 'cell' ? 'cell' : 'range'
      return { sheet, area: { kind, r, c } }
    },
  }
}

/** The source of a copy turned at `target`'s top-left cell on a scratch sheet of its own, or Excel's refusal of a turn off the sheet. */
export function turned(book: BookInput, clip: Clipboard, to: string, target: Rect): Clipboard | Refused {
  const { source, from, src } = clip
  const at = { r1: target.r1, c1: target.c1, r2: target.r1 + src.c2 - src.c1, c2: target.c1 + src.r2 - src.r1 }
  if (at.r2 >= MAX_ROWS || at.c2 >= MAX_COLS) return { ok: false, code: 'off_sheet', message: say(book.locale, { key: 'copyOffSheet', sheet: to, cell: formatAddress(target.r1, target.c1) }) }
  const turn = (r: number, c: number) => formatAddress(at.r1 + c - src.c1, at.c1 + r - src.r1)
  const turnRect = (x: Rect): Rect => ({ r1: at.r1 + x.c1 - src.c1, r2: at.r1 + x.c2 - src.c1, c1: at.c1 + x.r1 - src.r1, c2: at.c1 + x.r2 - src.r1 })
  const t = turnTransform(from, src, { r: at.r1, c: at.c1 })
  const names = sheetNames(source)
  const formula = (text: string) => rewriteFormula(text, from, from, t, names)
  // the looks of the cells without an entry as entries, so that each cell turns with its own
  const sheet = expanded(sheetOf(source, from), [src])
  const [rows, cols] = [lines(sheet.rows, 'r'), lines(sheet.columns, 'c')]
  const lined = rows.length > 0 || cols.length > 0
  const cells: Record<string, CellInput> = {}
  const g = gridOf(sheet)
  for (const i of cellsIn(g, src)) {
    spend(1)
    const [r, c, cell] = [g.row[i], g.col[i], g.cell[i]]
    const line = lined ? lookOf(undefined, entryAt(rows, r), entryAt(cols, c)) : {}
    const out: CellInput = { ...cell }
    // the scratch sheet has no lines: a cell shows the look its lines gave it
    if (out.style === undefined && line.style !== undefined) out.style = line.style
    if (out.numberFormat === undefined && line.numberFormat !== undefined) out.numberFormat = line.numberFormat
    if (cell.formula !== undefined) out.formula = formula(cell.formula)
    if (cell.array !== undefined) out.array = cornersText(turnRect(parseRange(cell.array)!))
    cells[turn(r, c)] = out
  }
  // the cells with neither an entry nor a look of the sheet's own that their lines give one
  if (lined)
    for (let r = src.r1; r <= src.r2; r++)
      for (let c = src.c1; c <= src.c2; c++) {
        spend(1)
        const a = turn(r, c)
        if (cells[a]) continue
        const line = lookOf(undefined, entryAt(rows, r), entryAt(cols, c))
        if (line.style !== undefined || line.numberFormat !== undefined) cells[a] = { ...(line.style !== undefined ? { style: line.style } : {}), ...(line.numberFormat !== undefined ? { numberFormat: line.numberFormat } : {}) }
      }
  const merges = mergesOf(sheet).flatMap((m) => inside([m], src).filter((x) => cellCount(x) > 1).map((x) => cornersText(turnRect(x))))
  // a rule's part over the source, its formulas made for that part's corner, then turned
  const ruleParts = (rects: Rect[], corner: { r: number; c: number }, made: (pieces: Rect[], text: (f: string) => string) => void) => {
    const pieces = inside(rects, src)
    if (pieces.length) made(pieces, (f) => formula(formulaAt(f, from, corner, cornerOf(pieces), names)))
  }
  const rules: ConditionalFormat[] = []
  for (const rule of sheet.conditionalFormats ?? []) ruleParts(rectsOf(rule), originOf(rule), (pieces, text) => rules.push({ ...mapFormulas(rule, text), ranges: pieces.map((x) => rectText(turnRect(x))) }))
  const validations: DataValidation[] = []
  for (const v of sheet.validations ?? []) ruleParts(areasOf(v.range)!, anchorOf(v), (pieces, text) => validations.push({ ...withFormulas(v, text), range: rangeText(pieces.map(turnRect)) }))
  const scratch: SheetInput = {
    name: TURNED,
    cells,
    ...(merges.length ? { merges } : {}),
    ...(rules.length ? { conditionalFormats: rules } : {}),
    ...(validations.length ? { validations } : {}),
  }
  return { source: { ...source, sheets: [...source.sheets, scratch] }, from: TURNED, src: at, packed: clip.packed }
}

/**
 * A transposed copy of the clipboard `held` pasted over `target` of sheet `to`, each tile turned at its own corner (the
 * paste set `held`: a reference with a `$` out of the source stays in every tile), and the area the tiles cover; no area
 * when the turn is refused.
 */
export function pasteTurned(book: BookInput, held: Clipboard, to: string, target: Rect, withMeta: boolean, looksOnly: boolean): { pasted: Applied; area: Rect | null } {
  const first = turned(book, held, to, target)
  if ('ok' in first) return { pasted: first, area: null }
  const { offsets } = tiles(first.src, target)
  const tile = ([dr, dc]: [number, number]) => (b: BookInput) => {
    const at = shiftRect(first.src, dr, dc)
    const clip = turned(b, held, to, at)
    return 'ok' in clip ? clip : pasteClip(b, clip, to, at, withMeta, looksOnly)
  }
  const pasted = offsets.length === 1 ? pasteClip(book, first, to, target, withMeta, looksOnly) : chain(book, offsets.map(tile))
  return { pasted, area: pasteArea(book, first, to, target) }
}

/** Why Excel refuses to paste contents onto a merged area of the target that the copy does not bring there (the paste set `probe`). */
export function contentRefusal(book: BookInput, full: BookInput, to: string, area: Rect): Refused | null {
  const ti = sheetIndex(book, to)
  const brought = new Set(full.sheets[ti].merges ?? [])
  const m = mergesOf(book.sheets[ti]).find((x) => meets(x, area) && !brought.has(cornersText(x)))
  return m ? { ok: false, code: 'merged', message: say(book.locale, { key: 'contentOnMerge', sheet: to, area: cornersText(m) }) } : null
}

/** The look fields of an entry, else of the sheet's looks at the cell. */
function lookAt(sheet: SheetInput, cell: CellInput | undefined, r: number, c: number): Pick<CellInput, 'style' | 'numberFormat'> {
  const x = cell ?? lookIn(sheet, r, c)
  return { style: x?.style, numberFormat: x?.numberFormat }
}

/**
 * The cell of the target after a paste of `parts`: its contents from the copy's cell or the target's, its look likewise
 * (of the entry, or of the sheet's looks where there is none), its note and hyperlink the target's; undefined for none.
 */
function laid(target: SheetInput, copy: SheetInput, a: string | undefined, b: string | undefined, r: number, c: number, parts: Parts): CellInput | undefined {
  const [t, x] = [a === undefined ? undefined : target.cells[a], b === undefined ? undefined : copy.cells[b]]
  const out: Record<string, unknown> = {}
  const has = parts.content ? x : t
  for (const k of CONTENT) if (has?.[k] !== undefined) out[k] = has[k]
  if (t?.note !== undefined) out.note = t.note
  if (t?.link !== undefined) out.link = t.link
  const look = parts.looks ? lookAt(copy, x, r, c) : lookAt(target, t, r, c)
  if (look.style !== undefined) out.style = look.style
  if (look.numberFormat !== undefined) out.numberFormat = look.numberFormat
  return Object.keys(out).length ? (out as CellInput) : undefined
}

/**
 * The book after a paste of `kind` (not `all`) over `area` of sheet `to`: `full` is the book after the whole copy, of
 * values for a paste of values. Its cells in the area are laid from both (`laid`); a paste of looks takes the copy's
 * sheet (its looks, lines, merged areas and conditional formats), the others the target's; data validation from the
 * copy for a paste of validation alone; a paste of formulas takes the names and links the copy brought.
 */
export function pastedPart(book: BookInput, full: BookInput, to: string, area: Rect, kind: Exclude<PasteKind, 'all'>): BookInput {
  const parts = PARTS[kind]
  const ti = sheetIndex(book, to)
  const [target, copy] = [book.sheets[ti], full.sheets[ti]]
  const base = parts.looks ? copy : target
  const cells: Record<string, CellInput> = {}
  const [g, h] = [gridOf(target), gridOf(copy)]
  const taken = new Map<string, [number, number, string | undefined, string | undefined]>()
  for (const i of cellsIn(g, area)) taken.set(g.addr[i], [g.row[i], g.col[i], g.addr[i], undefined])
  for (const i of cellsIn(h, area)) {
    const was = taken.get(h.addr[i])
    taken.set(h.addr[i], [h.row[i], h.col[i], was?.[2], h.addr[i]])
  }
  // the cells outside the area as the base has them, in its order, then the area's
  const gb = parts.looks ? h : g
  for (let i = 0; i < gb.addr.length; i++) if (!taken.has(gb.addr[i])) cells[gb.addr[i]] = gb.cell[i]
  // where the base's entry is gone, no look of the sheet's own: the entry showed its lines' look, not the one kept under it
  const bare: [number, number, number][] = []
  for (const [addr, [r, c, a, b]] of taken) {
    spend(1)
    const cell = laid(target, copy, a, b, r, c, parts)
    if (cell) cells[addr] = cell
    else if ((parts.looks ? b : a) !== undefined) bare.push([r, c, NONE])
  }
  const layer = layerOf(base)
  let sheet = layer ? withLayer(base, cells, bare.length ? edited(layer, cellEdits(bare), new LookTable(layer.looks)) : layer) : { ...base, cells }
  if (parts.looks || parts.validation) sheet = withField(sheet, 'validations', (parts.validation ? copy : target).validations)
  const sheets = book.sheets.map((s, i) => (i === ti ? sheet : s))
  const books = kind === 'formulas' ? full : book
  return { ...books, styles: full.styles, sheets }
}

/** How a copy pastes (`StructureOp` `copyRange`): the kind of paste, turned, with `meta`, and whether it is Paste Special (`paste` given). */
export interface PasteOptions {
  kind: PasteKind
  transpose: boolean
  withMeta: boolean
  special: boolean
}

/** The merged area of sheet `to` a target of one cell lies in, or undefined. */
function mergedAt(book: BookInput, to: string, target: Rect): Rect | undefined {
  if (target.r1 !== target.r2 || target.c1 !== target.c2) return undefined
  return mergesOf(sheetOf(book, to)).find((m) => m.r1 <= target.r1 && target.r1 <= m.r2 && m.c1 <= target.c1 && target.c1 <= m.c2)
}

/**
 * Copy and paste of `src` of sheet `from` over `target` of `to`: the cells as pasted (before a header cell pasted over
 * takes its value as its column's name) and the result. Paste Special takes its target as Excel takes a selection (the
 * paste set `held`): a cell of a merged area stands for the area, so a copy is pasted at the area's top-left cell, and a
 * copy of one cell over the whole area, but for `all`, which pastes it into the area's top-left cell and keeps the area.
 * A paste that writes no contents lays its source over the area alone, cut to the area's size (the set `merges`).
 */
export function pasteCopy(book: BookInput, from: string, src: Rect, to: string, target: Rect, o: PasteOptions, values: ValuesIn): { pasted: Applied; r: Applied } {
  const merged = o.special ? mergedAt(book, to, target) : undefined
  const one = src.r1 === src.r2 && src.c1 === src.c2
  if (merged) {
    const corner = { r1: merged.r1, c1: merged.c1, r2: merged.r1, c2: merged.c1 }
    if (one && o.kind === 'all') return intoMerged(book, from, src, to, corner, merged, o, values)
    if (!writesContent(o.kind)) {
      // turned, the source's rows lie along the area's columns
      const [h, w] = o.transpose ? [src.c2 - src.c1, src.r2 - src.r1] : [src.r2 - src.r1, src.c2 - src.c1]
      const [mh, mw] = [Math.min(h, merged.r2 - merged.r1), Math.min(w, merged.c2 - merged.c1)]
      src = o.transpose ? { ...src, r2: src.r1 + mw, c2: src.c1 + mh } : { ...src, r2: src.r1 + mh, c2: src.c1 + mw }
      target = merged
    } else target = one ? merged : corner
  }
  const looksOnly = !writesContent(o.kind)
  // the source as the clipboard holds it: of values for a paste of values, turned for a transposed one
  const held = clipboard(book, from, src, values, ofValues(o.kind))
  const turn = o.transpose ? pasteTurned(book, held, to, target, o.withMeta, looksOnly) : null
  const pasted = turn ? turn.pasted : pasteClip(book, held, to, target, o.withMeta, looksOnly)
  const area = turn ? turn.area : pasteArea(book, held, to, target)
  // a transposed copy's cells are not where its source's tables would read them: it renames header cells alone
  let r = turn ? (area ? headersPasted(book, to, area)(pasted) : pasted) : copyTables(book, from, src, to, target, o.kind === 'all').after(pasted)
  if (o.kind !== 'all' && r.ok && area) {
    const refused = writesContent(o.kind) ? contentRefusal(book, r.book, to, area) : null
    const part: Applied = refused ?? { ...r, book: pastedPart(book, r.book, to, area, o.kind), geometry: writesContent(o.kind) ? r.geometry : [] }
    r = writesContent(o.kind) ? headersPasted(book, to, area)(part) : part
  }
  // the looks the copy laid, over the rows a session fits to their fonts
  if (r.ok && area && (o.kind === 'all' || o.kind === 'formats')) r = { ...r, looked: { s: sheetIndex(book, to), area } }
  return { pasted, r }
}

/** A copy of one cell pasted into the top-left cell (`corner`) of the merged area `merged`, which stays. */
function intoMerged(book: BookInput, from: string, src: Rect, to: string, corner: Rect, merged: Rect, o: PasteOptions, values: ValuesIn): { pasted: Applied; r: Applied } {
  const ti = sheetIndex(book, to)
  const text = cornersText(merged)
  const without = { ...book, sheets: book.sheets.map((s, i) => (i === ti ? withField(s, 'merges', s.merges?.filter((m) => cornersText(parseMerge(m)!) !== text)) : s)) }
  const kept = (x: Applied): Applied => (x.ok ? { ...x, book: { ...x.book, sheets: x.book.sheets.map((s, i) => (i === ti ? { ...s, merges: [...(s.merges ?? []), text] } : s)) } } : x)
  const { pasted, r } = pasteCopy(without, from, src, to, corner, { ...o, special: false }, values)
  return { pasted: kept(pasted), r: kept(r) }
}
