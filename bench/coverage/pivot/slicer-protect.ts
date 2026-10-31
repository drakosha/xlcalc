/**
 * The set `slicer-protect`: a slicer and a timeline on a protected sheet (`Worksheet.Protect` without a password) with
 * Edit objects allowed or not and Use PivotTable reports allowed or not, both locked (`Slicer.Locked`, as Excel adds
 * them) or not. On each sheet: an item chosen (`SlicerItem.Selected`), the timeline's dates set, each moved (`Top`) and
 * resized (`Width`); Excel's error where it refuses, and what each shows after. Predicted before Excel ran it:
 * `slicer-protect.predicted.json`.
 */
import { serial } from './data.js'
import type { PivotCase, PivotOp } from './cases.js'

const AFTER = ['Top', 'Width']

export function buildPivotCases(): PivotCase[] {
  const out: PivotCase[] = []
  for (const objects of [false, true])
    for (const pivots of [false, true])
      for (const locked of [true, false]) {
        const ops: PivotOp[] = [
          { rows: ['Region'] },
          { value: { field: 'Price' } },
          { slicer: 'Region', name: 'Slicer_Region' },
          { timeline: 'Date', name: 'Timeline_Date' },
          ...(locked ? [] : (['Slicer_Region', 'Timeline_Date'].map((n) => ({ slicerSet: n, prop: 'Locked', to: false })) as PivotOp[])),
          { protect: 'P', objects, pivots },
          { select: 'Slicer_Region', items: ['West'] },
          { timelineSet: 'Timeline_Date', from: serial(2024, 3, 1), to: serial(2024, 12, 31) },
          { slicerSet: 'Slicer_Region', prop: 'Top', to: 30 },
          { slicerSet: 'Slicer_Region', prop: 'Width', to: 180 },
          { slicerSet: 'Timeline_Date', prop: 'Top', to: 250 },
          { slicerSet: 'Timeline_Date', prop: 'Width', to: 300 },
          { slicerGet: 'Slicer_Region', get: AFTER },
          { slicerGet: 'Timeline_Date', get: AFTER },
          { unprotect: 'P' },
        ]
        out.push({ id: `${objects ? 'objects' : 'no-objects'}-${pivots ? 'pivots' : 'no-pivots'}-${locked ? 'locked' : 'unlocked'}`, ops })
      }
  return out
}
