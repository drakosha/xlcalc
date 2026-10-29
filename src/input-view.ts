/** Saved view metadata checked at the book boundary; it does not enter computation. */
import { InputError } from './input-error.js'
import { array, checkBoolean, checkNumber, checkOneOf, object, onlyKeys } from './input/basic.js'
import { checkAddress } from './input/range.js'

export function checkSheetView(v: unknown, path: string): void {
  const view = object(v, path, 'a sheet view such as { activeCell: "B2", zoom: 125 }')
  onlyKeys(view, path, ['activeCell', 'selection', 'topLeft', 'zoom', 'gridlines', 'headings', 'split'])
  for (const field of ['gridlines', 'headings']) if (view[field] !== undefined) checkBoolean(view[field], `${path}.${field}`)
  for (const field of ['activeCell', 'topLeft']) if (view[field] !== undefined) checkAddress(view[field], `${path}.${field}`)
  if (view.zoom !== undefined) checkNumber(view.zoom, `${path}.zoom`, 10, 400, true)
  if (view.selection !== undefined) {
    const ranges = array(view.selection, `${path}.selection`, 'an array of A1 cells or ranges')
    if (!ranges.length) throw new InputError(`${path}.selection`, 'expected at least one cell or range')
    ranges.forEach((range, i) => {
      const at = `${path}.selection[${i}]`
      if (typeof range !== 'string' || range.split(':').length > 2) throw new InputError(at, 'expected an A1 cell or range')
      for (const cell of range.split(':')) checkAddress(cell, at)
    })
  }
  if (view.split !== undefined) checkSplit(view.split, `${path}.split`)
}

/** The largest position of a split's line, in points: far past any screen. */
const MAX_SPLIT = 100_000

/** A split window: a line on at least one axis, and an active pane that lies by the lines it has. */
function checkSplit(v: unknown, path: string): void {
  const split = object(v, path, 'a split window such as { x: 120, y: 60, topLeft: "F20" }')
  onlyKeys(split, path, ['x', 'y', 'topLeft', 'activePane'])
  for (const axis of ['x', 'y']) if (split[axis] !== undefined) checkNumber(split[axis], `${path}.${axis}`, 0, MAX_SPLIT)
  if (!split.x && !split.y) throw new InputError(path, 'expected a line: x or y above 0')
  if (split.topLeft !== undefined) checkAddress(split.topLeft, `${path}.topLeft`)
  if (split.activePane !== undefined) checkOneOf(split.activePane, `${path}.activePane`, ['topLeft', ...(split.x ? ['topRight'] : []), ...(split.y ? ['bottomLeft'] : []), ...(split.x && split.y ? ['bottomRight'] : [])])
}
