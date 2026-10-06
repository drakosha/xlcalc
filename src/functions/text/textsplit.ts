/** TEXTSPLIT (Excel 365), this engine's own: text cut by column and row delimiters into an array. */
import type { Cx, FnDef, Segment } from '../../args.js'
import { Arr } from '../../arrays/value.js'
import { E, isErr, type V } from '../../values.js'
import { delimiters } from './common.js'

/** The earliest delimiter at or after `from` (a column delimiter wins a tie, then the longer). */
function next(text: string, from: number, cols: string[], rows: string[]): { at: number; len: number; row: boolean } | null {
  let best: { at: number; len: number; row: boolean } | null = null
  for (const [keys, row] of [[cols, false], [rows, true]] as const)
    for (const k of keys) {
      const at = text.indexOf(k, from)
      if (at < 0) continue
      if (!best || at < best.at || (at === best.at && !row && best.row) || (at === best.at && row === best.row && k.length > best.len)) best = { at, len: k.length, row }
    }
  return best
}

function textsplit(cx: Cx, text: string, colSeg: Segment, rowSeg?: Segment, ignoreEmpty = false, mode = 0, pad: V = E.NA): V | Arr {
  const cols = delimiters(colSeg, cx.rules)
  if (isErr(cols)) return cols
  const rows = rowSeg === undefined ? [] : delimiters(rowSeg, cx.rules)
  if (isErr(rows)) return rows
  if (mode !== 0 && mode !== 1) return E.VALUE
  // an empty delimiter anywhere (a blank cell too) is #VALUE!, as in Excel
  if (cols.includes('') || rows.includes('') || text === '') return E.VALUE
  const hay = mode ? text.toLowerCase() : text
  const fold = (ks: string[]) => (mode ? ks.map((k) => k.toLowerCase()) : ks)
  const lines: string[][] = [[]]
  let p = 0
  for (;;) {
    const d = next(hay, p, fold(cols), fold(rows))
    lines.at(-1)!.push(text.slice(p, d ? d.at : text.length))
    if (!d) break
    if (d.row) lines.push([])
    p = d.at + d.len
  }
  const kept = (ignoreEmpty ? lines.map((l) => l.filter((s) => s !== '')).filter((l) => l.length) : lines)
  if (!kept.length) return E.VALUE
  const w = Math.max(...kept.map((l) => l.length))
  const v: V[] = []
  for (const l of kept) for (let j = 0; j < w; j++) v.push(j < l.length ? l[j] : pad)
  return new Arr(kept.length, w, kept.length, w, v)
}

export const functions: Record<string, FnDef> = {
  TEXTSPLIT: { args: 's V V? b? i? a?', array: true, cx: true, fn: textsplit },
}
