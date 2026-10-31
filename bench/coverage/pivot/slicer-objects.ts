/**
 * The set `slicer-objects`: a slicer and a timeline as objects of the sheet. Renamed (`Slicer.Name`): to a free name,
 * one taken by another slicer or a timeline, in another case, empty, long, with odd characters, to its own cache's
 * name or another's, with a caption set first; what becomes of the caption. `Locked` and `DisableMoveResizeUI` set off
 * and on, for the file Excel saves. Where Excel puts new ones without a place (`Slicers.Add` without its position) and
 * how large: one, three, slicers and timelines in turn, on the PivotTable's sheet and on another. Each read through
 * `slicerGet`. Predicted before Excel ran it: `slicer-objects.predicted.json`.
 */
import type { PivotCase, PivotOp } from './cases.js'

const OWN = ['Name', 'Caption', 'SlicerCache.Name', 'Locked', 'DisableMoveResizeUI']
const PLACE = ['Top', 'Left', 'Width', 'Height', 'Shape.TopLeftCell.Row', 'Shape.TopLeftCell.Column', 'Shape.BottomRightCell.Row', 'Shape.BottomRightCell.Column']

const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Price' } }]
type Made = Extract<PivotOp, { name: string }>
const region: Made = { slicer: 'Region', name: 'Slicer_Region' }
const product: Made = { slicer: 'Product', name: 'Slicer_Product' }
const date: Made = { timeline: 'Date', name: 'Timeline_Date' }
const get = (name: string, more: string[] = []): PivotOp => ({ slicerGet: name, get: [...OWN, ...more] })
const rename = (name: string, to: string): PivotOp => ({ slicerSet: name, prop: 'Name', to })

export function buildPivotCases(): PivotCase[] {
  const auto = (o: Made, on?: string): PivotOp[] => [{ ...o, auto: true, ...(on ? { on } : {}) }, { slicerGet: o.name, get: PLACE }]
  return [
    { id: 'rename', ops: [...base, region, rename('Slicer_Region', 'Regions'), get('Regions')] },
    { id: 'rename-taken', ops: [...base, region, product, rename('Slicer_Product', 'Slicer_Region'), get('Slicer_Product'), get('Slicer_Region')] },
    { id: 'rename-case', ops: [...base, region, product, rename('Slicer_Product', 'SLICER_REGION'), get('Slicer_Product')] },
    { id: 'rename-timeline', ops: [...base, region, date, rename('Slicer_Region', 'Timeline_Date'), get('Slicer_Region'), rename('Timeline_Date', 'Dates'), get('Dates')] },
    { id: 'rename-empty', ops: [...base, region, rename('Slicer_Region', ''), get('Slicer_Region'), rename('Slicer_Region', ' '), get('Slicer_Region'), get(' ')] },
    { id: 'rename-long', ops: [...base, region, rename('Slicer_Region', 'L'.repeat(255)), get('L'.repeat(255)), rename('L'.repeat(255), 'M'.repeat(256)), get('L'.repeat(255)), get('M'.repeat(256))] },
    { id: 'rename-chars', ops: [...base, region, rename('Slicer_Region', "A/B [x]*?:'"), get("A/B [x]*?:'")] },
    { id: 'rename-cache', ops: [...base, region, product, rename('Slicer_Region', 'R'), rename('Slicer_Product', 'Slicer_Region'), get('Slicer_Region'), rename('R', 'Slicer_Product'), get('Slicer_Product')] },
    { id: 'rename-caption', ops: [...base, region, { slicerSet: 'Slicer_Region', prop: 'Caption', to: 'Pick a region' }, rename('Slicer_Region', 'Regions'), get('Regions')] },
    { id: 'rename-range-name', ops: [...base, region, rename('Slicer_Region', 'A1'), get('A1'), rename('A1', 'Total'), get('Total')] },
    { id: 'locked-off', ops: [...base, region, date, { slicerSet: 'Slicer_Region', prop: 'Locked', to: false }, { slicerSet: 'Timeline_Date', prop: 'Locked', to: false }, get('Slicer_Region'), get('Timeline_Date')] },
    { id: 'move-off', ops: [...base, region, date, { slicerSet: 'Slicer_Region', prop: 'DisableMoveResizeUI', to: true }, { slicerSet: 'Timeline_Date', prop: 'DisableMoveResizeUI', to: true }, get('Slicer_Region'), get('Timeline_Date')] },
    { id: 'place-one', ops: [...base, ...auto(region)] },
    { id: 'place-timeline', ops: [...base, ...auto(date)] },
    { id: 'place-three', ops: [...base, ...auto(region), ...auto(product), ...auto({ slicer: 'Flag', name: 'Slicer_Flag' })] },
    { id: 'place-mixed', ops: [...base, ...auto(region), ...auto(date), ...auto(product), ...auto({ timeline: 'When', name: 'Timeline_When' })] },
    { id: 'place-other-sheet', ops: [...base, ...auto(region, 'G'), ...auto(date, 'G'), ...auto(product, 'G')] },
  ]
}
