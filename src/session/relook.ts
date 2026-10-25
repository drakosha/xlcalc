/**
 * A session's sheets given new looks whole (`Plan.layers`), and the step of the history that puts them back
 * (`undo.ts`): the cells whose look changed listed with their style and number format, a rectangle of them larger than
 * LISTED_LOOKS given whole as a range to read again (`EditResult.repaint`), as after a structure operation. The way back
 * holds the rectangles that changed and their looks before, so that it changes those alone: the cells the same call
 * writes one by one keep their looks, which their own step puts back.
 */
import type { LookRange } from '../contract.js'
import type { EngineState } from '../engine.js'
import { cellEdits, edited, emptyLayer, LookTable, rectEdits, type Layer, type RowsEdit } from '../looks/layer.js'
import { clearing, differing } from '../looks/move.js'
import { layerOf, setLayer } from '../looks/sheet.js'
import { formatAddress, placeOf } from '../model.js'
import { areaText, cellCount, parseRange, type Rect } from '../structure/area.js'
import { fieldsBetween, nothing, type Outcome } from './answer.js'
import { LISTED_LOOKS } from './moved.js'

/** The looks of a sheet's rectangles (`areas`, apart) as a step holds them: each piece of one look as a range of the looks. */
export interface LooksBack {
  i: number
  areas: string[]
  looks: LookRange[]
}

/** The looks of `layer` in rectangles `xs`, a range of one look each. */
function looksIn(layer: Layer | undefined, xs: Rect[]): LookRange[] {
  const out: LookRange[] = []
  for (const x of xs) for (const p of layer?.pieces(x) ?? []) out.push({ range: areaText(p), ...layer!.looks.list[p.look] })
  return out
}

/** Rectangles apart, each of one look of `looks`, as edits of the rows they cover. */
const laying = (xs: LookRange[], looks: LookTable): RowsEdit[] => rectEdits(xs.map((x) => ({ ...parseRange(x.range)!, look: looks.of(x) })))

/**
 * The sheets' looks set whole, each `[sheet, layer]`; `kept`: cells (`[sheet, row, column]`) whose looks stay as they
 * are. What changed, the ways back and the looks before.
 */
export function relook(s: EngineState, layers: Iterable<[number, Layer | undefined]>, kept: [number, number, number][] = []): [Outcome, LooksBack[], [number, Layer | undefined][]] {
  const out = nothing()
  const [back, was]: [LooksBack[], [number, Layer | undefined][]] = [[], []]
  for (const [i, given] of layers) {
    const sheet = s.input.sheets[i]
    const old = layerOf(sheet)
    let layer = given
    const stay = kept.filter(([j]) => j === i)
    if (stay.length) {
      const looks = new LookTable(layer?.looks)
      layer = edited(layer ?? emptyLayer(), cellEdits(stay.map(([, r, c]) => [r, c, looks.of(old?.lookAt(r, c))])), looks)
    }
    const xs = [...differing(old, layer)]
    if (!xs.length) continue
    back.push({ i, areas: xs.map(areaText), looks: looksIn(old, xs) })
    was.push([i, old])
    for (const x of xs) {
      if (cellCount(x) > LISTED_LOOKS) {
        out.looks.push({ sheet: s.sheetNames[i], range: areaText(x) })
        continue
      }
      for (let r = x.r1; r <= x.r2; r++)
        for (let c = x.c1; c <= x.c2; c++) {
          // a cell of an entry shows no look of the sheet's looks
          if (sheet.cells[formatAddress(r, c)]) continue
          const f = fieldsBetween(old?.lookAt(r, c), undefined, layer?.lookAt(r, c), undefined)
          if (!f) continue
          out.fields.set(placeOf(i, r, c), f)
          out.keys.push(placeOf(i, r, c))
        }
    }
    setLayer(sheet, layer ?? emptyLayer())
  }
  out.keys.sort((a, b) => a - b)
  return [out, back, was]
}

/** The ways back (`relook`) taken over the sheets' looks as they are: what changed and the ways forth. */
export function relookBack(s: EngineState, backs: LooksBack[]): [Outcome, LooksBack[]] {
  const layers = backs.map(({ i, areas, looks }): [number, Layer] => {
    const now = layerOf(s.input.sheets[i]) ?? emptyLayer()
    const table = new LookTable(now.looks)
    return [i, edited(edited(now, clearing(areas.map((a) => parseRange(a)!))), laying(looks, table), table)]
  })
  const [outcome, back] = relook(s, layers)
  return [outcome, back]
}
