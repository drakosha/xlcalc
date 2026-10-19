/** Saved view metadata checked at the book boundary; it does not enter computation. */
import { InputError } from './input-error.js'
import { array, checkNumber, object, onlyKeys } from './input/basic.js'
import { checkAddress } from './input/range.js'

export function checkSheetView(v: unknown, path: string): void {
  const view = object(v, path, 'a sheet view such as { activeCell: "B2", zoom: 125 }')
  onlyKeys(view, path, ['activeCell', 'selection', 'topLeft', 'zoom'])
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
}
