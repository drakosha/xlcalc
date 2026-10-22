/** Ordered operations staged over sparse cell overlays; no full snapshot is made for ordinary cell and look batches. */
import type { BookInput, CellInput, CellReplace, SessionOperation, SheetInput } from '../contract.js'
import type { EngineState } from '../engine.js'
import { allotWrites } from '../exec/budget.js'
import { checkEdits, checkLayoutEdits, checkNames } from '../input.js'
import { InputError } from '../input-error.js'
import { array, isObject, object, onlyKeys } from '../input/basic.js'
import { checkStyles } from '../input/styles.js'
import { sheetKey } from '../syntax.js'
import type { Layer } from '../looks/layer.js'
import { withLayer, withLooksOf, withoutLooksAt } from '../looks/sheet.js'
import { parseAddress } from '../model.js'
import { linksIn, markedLinks } from '../syntax/links.js'
import { editsTaken } from '../taken.js'
import { hiddenRows } from '../layout/hidden.js'
import { checkArrayEdits } from '../spill/fixed.js'
import { headerEdits } from '../tables/headers.js'
import { checkTables } from '../tables/check.js'
import { checkStylesUsed, applyLayout, inputAt } from './book.js'
import { checkFieldEdits, resolveEdits } from './cells.js'
import { clearPlan } from './clear.js'
import { checkPatches } from './patch.js'
import { patchPlan, type Plan } from './styles.js'

export interface OperationsPlan {
  book: BookInput
  edits: CellReplace[]
  back: CellReplace[]
  rebuild: boolean
  layout: Map<number, SheetInput>
  /** The sheets' looks set whole by patches (`Plan.layers`), laid before the edits. */
  layers: Map<number, Layer>
  styles: boolean
}

/** Stages all input checks before the live state can be changed. Cell records are copied only for a rebuild. */
export function operationsPlan(s: EngineState, given: SessionOperation[], headerText: Parameters<typeof headerEdits>[2]): OperationsPlan {
  const book: BookInput = { ...s.input, sheets: s.input.sheets.slice() }
  const edits = new Map<string, CellReplace>(),
    back = new Map<string, CellReplace>()
  const overlays = new Map<number, Map<string, CellInput | null>>()
  const layers = new Map<number, Layer>()
  let rebuild = false,
    styles = false
  const index = (name: string) => s.sheetIdx.get(sheetKey(name))
  const draft = () => ({ ...s, input: book })
  const cells = (i: number) => {
    let changes = overlays.get(i)
    if (changes) return changes
    changes = new Map()
    overlays.set(i, changes)
    const base = book.sheets[i].cells,
      own = changes
    const proxy = new Proxy(base, {
      get: (target, k) => (typeof k === 'string' && own.has(k) ? (own.get(k) ?? undefined) : Reflect.get(target, k)),
      has: (target, k) => (typeof k === 'string' && own.has(k) ? own.get(k) !== null : Reflect.has(target, k)),
      ownKeys: (target) => [...new Set([...Reflect.ownKeys(target), ...own.keys()])].filter((k) => typeof k !== 'string' || !own.has(k) || own.get(k) !== null),
      getOwnPropertyDescriptor: (target, k) => (typeof k === 'string' && own.has(k) ? (own.get(k) === null ? undefined : { configurable: true, enumerable: true, writable: true, value: own.get(k) }) : Reflect.getOwnPropertyDescriptor(target, k)),
    })
    book.sheets[i] = { ...book.sheets[i], cells: proxy }
    return changes
  }
  const put = (items: CellReplace[]) => {
    // the cells left without an input, which the sheets' looks of the draft give none either
    const gone = new Map<number, [number, number][]>()
    for (const e of items) {
      const i = index(e.sheet)!,
        k = `${i}:${e.cell}`
      if (!back.has(k)) back.set(k, { sheet: s.sheetNames[i], cell: e.cell, input: inputAt(s, i, e.cell) ?? null })
      cells(i).set(e.cell, e.input)
      edits.set(k, { ...e, sheet: s.sheetNames[i] })
      if (e.input === null) {
        const { row, col } = parseAddress(e.cell)
        ;(gone.get(i) ?? gone.set(i, []).get(i)!).push([row, col])
      }
    }
    for (const [i, at] of gone) {
      book.sheets[i] = withoutLooksAt(book.sheets[i], at)
    }
  }
  /** A plan's looks set whole over the draft's sheets, as they are laid before its cells. */
  const relaid = (p: Plan) =>
    p.layers?.size ? { ...book, sheets: book.sheets.map((sh, i) => (p.layers!.has(i) ? withLayer(sh, sh.cells, p.layers!.get(i)!) : sh)) } : book
  /** A plan staged: its sheets' lines and rules over the draft's sheets, which keep their looks, then its looks set whole, then its cells. */
  const looks = (p: Plan) => {
    for (const [i, sh] of p.sheets) book.sheets[i] = withLooksOf({ ...sh, cells: book.sheets[i].cells }, book.sheets[i])
    for (const [i, layer] of p.layers ?? []) {
      book.sheets[i] = withLayer(book.sheets[i], book.sheets[i].cells, layer)
      layers.set(i, layer)
      // a cell an operation before left without an entry (its edit is laid after the looks) takes the look these give it
      const own = overlays.get(i)
      for (const [k, e] of edits) {
        if (index(e.sheet) !== i || own?.get(e.cell) !== null) continue
        const { row, col } = parseAddress(e.cell)
        const look = layer.lookAt(row, col)
        edits.set(k, { ...e, input: look ? { ...look } : null })
      }
    }
    put(p.edits)
    if (p.styles) {
      book.styles = p.styles
      styles = true
    }
  }
  const ops = array(given, 'operations', 'an array of serializable session operations')
  // the batch writes the cells of its edits: over `limits.writes` it stops before anything is staged
  allotWrites(() => ops.reduce((n: number, op) => n + (isObject(op) && (op.op === 'setCells' || op.op === 'replaceCells') && Array.isArray(op.edits) ? op.edits.length : 0), 0), s.input.locale)
  for (const [n, value] of ops.entries()) {
    const path = `operations[${n}]`,
      op = object(value, path, 'a session operation')
    const count = book.styles?.length ?? 0
    switch (op.op) {
      case 'setCells':
      case 'replaceCells': {
        onlyKeys(op, path, ['op', 'edits'])
        const items = op.op === 'setCells' ? resolveEdits(draft(), editsTaken(checkFieldEdits(op.edits, index, count, `${path}.edits`))) : (checkEdits(op.edits, index, count, `${path}.edits`) as CellReplace[])
        const renamed = headerEdits(book, items, headerText)
        if (renamed) {
          Object.assign(book, renamed)
          rebuild = true
          overlays.clear()
        } else put(items)
        break
      }
      case 'clearCells': {
        onlyKeys(op, path, ['op', 'ranges', 'what'])
        const p = clearPlan(draft(), op.ranges as never, op.what as never)
        // a header cell cleared renames its column, as `Session.clearCells` does: the plan's sheets go in first, so the rename rewrites their rules too
        const looked = relaid(p)
        const laid = p.sheets.size ? { ...looked, sheets: looked.sheets.map((sh, i) => (p.sheets.has(i) ? withLooksOf({ ...p.sheets.get(i)!, cells: sh.cells }, sh) : sh)) } : looked
        const renamed = headerEdits(laid, p.edits, headerText)
        if (!renamed) {
          looks(p)
          break
        }
        Object.assign(book, renamed)
        if (p.styles) {
          book.styles = p.styles
          styles = true
        }
        rebuild = true
        overlays.clear()
        break
      }
      case 'patchCells':
        onlyKeys(op, path, ['op', 'patches'])
        looks(patchPlan(draft(), checkPatches(op.patches, index, count)))
        break
      case 'setNames':
        onlyKeys(op, path, ['op', 'names'])
        book.names = checkNames(op.names, `${path}.names`)
        rebuild = true
        break
      case 'setStyles':
        onlyKeys(op, path, ['op', 'styles'])
        book.styles = checkStyles(op.styles, `${path}.styles`)
        styles = true
        break
      case 'setLayout': {
        onlyKeys(op, path, ['op', 'edits'])
        const checked = checkLayoutEdits(op.edits, index, count, `${path}.edits`)
        if (checked.some((e) => 'fit' in e)) throw new InputError(path, 'AutoFit reads calculated values; use setLayout after applyChanges')
        applyLayout(book, checked as never, (name) => index(name)!)
        break
      }
      default:
        throw new InputError(`${path}.op`, 'expected setCells, replaceCells, clearCells, patchCells, setNames, setStyles or setLayout')
    }
  }
  checkTables(book, 'book')
  if (styles) checkStylesUsed(book, book.styles?.length ?? 0)
  const final = [...edits.values()]
  if (book.externalBooks?.length) {
    const entered = new Set(
      final.flatMap((e) => {
        const old = s.input.sheets[index(e.sheet)!].cells[e.cell]
        return e.input?.formula && JSON.stringify([e.input.formula, e.input.array]) !== JSON.stringify([old?.formula, old?.array]) ? linksIn(e.input.formula) : []
      }),
    )
    if (entered.size) {
      book.externalBooks = markedLinks(book.externalBooks, entered)
      rebuild = true
    }
  }
  checkArrayEdits(s, final)
  const layout = new Map<number, SheetInput>()
  book.sheets.forEach((sh, i) => {
    const old = s.input.sheets[i]
    const { cells: _a, looks: _c, ...a } = old,
      { cells: _b, looks: _d, ...b } = sh
    if (JSON.stringify(a) !== JSON.stringify(b)) layout.set(i, sh)
    if (JSON.stringify(old.merges ?? []) !== JSON.stringify(sh.merges ?? []) || JSON.stringify(hiddenRows(old) ?? null) !== JSON.stringify(hiddenRows(sh) ?? null) || old.filtered !== sh.filtered) rebuild = true
  })
  if (rebuild) book.sheets = book.sheets.map((sh) => ({ ...sh, cells: { ...sh.cells } }))
  return { book, edits: final, back: [...back.values()], rebuild, layout, layers, styles }
}
