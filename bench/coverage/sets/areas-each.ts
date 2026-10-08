/**
 * SUBTOTAL, AGGREGATE, FREQUENCY and SORTBY given a union (`(D!B1:B3,D!C1:C3)`) or a reference across sheets
 * (`Jan:Mar!A1:A3`): each area taken in turn, nested SUBTOTAL cells skipped inside a union, errors, overlaps, and the
 * array answers read through ARRAYTOTEXT. The set `areas` found that Excel takes these references there; this set
 * measures how. Answers: `test/excel/cases/areas-each.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { SHEETS as AREA_SHEETS } from './areas.js'

/** Jan, Feb and Mar of the set `areas`, and T: values with a SUBTOTAL among them (A3), which SUBTOTAL skips. */
export const SHEETS: SheetCells = { ...AREA_SHEETS, T: { A1: 1, A2: 2, A3: { formula: 'SUBTOTAL(9,A1:A2)' }, B1: 10 } }

const n = (v: number): Val => ({ t: 'n', v })
const s = (v: string): Val => ({ t: 's', v })
const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })

/** Formula, function and the answer predicted before Excel ran (undefined: not predicted). */
const CASES: [string, string, Val?][] = [
  // SUBTOTAL over a union: each area's values together (D!B1:B3 4, 8, 15; D!C1:C3 2, 7, 11)
  ['SUBTOTAL(9,(D!B1:B3,D!C1:C3))', 'SUBTOTAL', n(47)],
  ['SUBTOTAL(1,(D!B1:B3,D!C1:C3))', 'SUBTOTAL', n(47 / 6)],
  ['SUBTOTAL(2,(D!B1:B3,D!D1:D3))', 'SUBTOTAL', n(4)],
  ['SUBTOTAL(3,(D!B1:B3,D!D1:D5))', 'SUBTOTAL', n(7)],
  ['SUBTOTAL(4,(D!B1:B3,D!C4:C6))', 'SUBTOTAL', n(40)],
  ['SUBTOTAL(5,(D!B1:B3,D!C4:C6))', 'SUBTOTAL', n(4)],
  ['SUBTOTAL(6,(D!B1:B2,D!C1:C2))', 'SUBTOTAL', n(448)],
  ['SUBTOTAL(7,(D!B1:B3,D!C1:C3))', 'SUBTOTAL'],
  ['SUBTOTAL(10,(D!B1:B3,D!C1:C3))', 'SUBTOTAL'],
  ['SUBTOTAL(109,(D!B1:B3,D!C1:C3))', 'SUBTOTAL', n(47)],
  ['SUBTOTAL(9,(D!B1:B3,D!B2:B4))', 'SUBTOTAL', n(66)],
  ['SUBTOTAL(9,(D!B1:B2,D!E1:E3))', 'SUBTOTAL', e('DIV0')],
  ['SUBTOTAL(9,(D!B1,D!C1,D!B8))', 'SUBTOTAL', n(16.25)],
  ['SUBTOTAL(9,D!B1,(D!C1,D!C2))', 'SUBTOTAL', n(13)],
  ['SUBTOTAL(2,(D!B:B,D!C1))', 'SUBTOTAL', n(9)],
  ['SUBTOTAL(9,(T!A1:A3,T!B1))', 'SUBTOTAL', n(13)],
  ['SUBTOTAL(9,(D!B1,Jan!A1))', 'SUBTOTAL', e('VALUE')],
  ['SUBTOTAL(9,Jan:Mar!A1)', 'SUBTOTAL', e('VALUE')],
  ['SUBTOTAL(9,(D!B1:B3,D!C1:C3),Jan:Mar!A1)', 'SUBTOTAL', e('VALUE')],
  // AGGREGATE: options 0–7 over a union, the array form (14–19) with a union as the array and as k
  ['AGGREGATE(9,4,(D!B1:B3,D!C1:C3))', 'AGGREGATE', n(47)],
  ['AGGREGATE(9,6,(D!B1:B2,D!E1:E3))', 'AGGREGATE', n(15)],
  ['AGGREGATE(9,4,(D!B1:B2,D!E1:E3))', 'AGGREGATE', e('DIV0')],
  ['AGGREGATE(4,0,(D!B1:B3,D!C4:C6))', 'AGGREGATE', n(40)],
  ['AGGREGATE(12,6,(D!B1:B3,D!C1:C3))', 'AGGREGATE', n(7.5)],
  ['AGGREGATE(9,0,(T!A1:A3,T!B1))', 'AGGREGATE', n(13)],
  ['AGGREGATE(9,4,(T!A1:A3,T!B1))', 'AGGREGATE', n(16)],
  ['AGGREGATE(14,4,(D!B1:B3,D!C1:C3),2)', 'AGGREGATE', n(11)],
  ['AGGREGATE(15,6,(D!B1:B2,D!E1:E3),1)', 'AGGREGATE', n(1)],
  ['AGGREGATE(14,6,D!B1:B8,(D!AH2,D!AH3))', 'AGGREGATE'],
  ['AGGREGATE(9,4,Jan:Mar!A1:A3)', 'AGGREGATE', e('VALUE')],
  ['AGGREGATE(14,4,Jan:Mar!A1:A3,1)', 'AGGREGATE'],
  ['AGGREGATE(9,4,D!B1,(D!C1,D!C2))', 'AGGREGATE', n(13)],
  // FREQUENCY: the data or the bins a union or a reference across sheets (bins D!AT1:AT3 0, 10, 20)
  ['ARRAYTOTEXT(FREQUENCY((D!B1:B3,D!C1:C3),D!AT1:AT3),1)', 'FREQUENCY', s('{0;4;2;0}')],
  ['ARRAYTOTEXT(FREQUENCY(Jan:Mar!A1:A3,D!AT1:AT3),1)', 'FREQUENCY', s('{0;9;0;0}')],
  ['ARRAYTOTEXT(FREQUENCY(D!B1:B8,(D!AT1,D!AT3)),1)', 'FREQUENCY', s('{1;5;2}')],
  ['ARRAYTOTEXT(FREQUENCY(D!B1:B8,Jan:Mar!A1),1)', 'FREQUENCY', s('{1;1;0;6}')],
  ['ROWS(FREQUENCY((D!B1:B3,D!C1:C3),D!AT1:AT3))', 'FREQUENCY', n(4)],
  ['ARRAYTOTEXT(FREQUENCY((D!B1,D!C1),5),1)', 'FREQUENCY', s('{2;0}')],
  ['ARRAYTOTEXT(FREQUENCY((D!D1:D4,D!B1),D!AT1:AT3),1)', 'FREQUENCY', s('{0;2;0;0}')],
  ['ARRAYTOTEXT(FREQUENCY((D!E1:E4,D!B1),D!AT1:AT3),1)', 'FREQUENCY'],
  ['ARRAYTOTEXT(FREQUENCY((D!B1:B3,Jan!A1),D!AT1:AT3),1)', 'FREQUENCY', e('VALUE')],
  ['ARRAYTOTEXT(FREQUENCY(Jan:Mar!B1:B3,D!AT1:AT3),1)', 'FREQUENCY', s('{0;1;1;4}')],
  // SORTBY: the array or a by-array a union or a reference across sheets, the areas taken one after another as VSTACK does
  ['ARRAYTOTEXT(SORTBY((D!B1:B3,D!C1:C3),{6;5;4;3;2;1}),1)', 'SORTBY', s('{11;7;2;15;8;4}')],
  ['ARRAYTOTEXT(SORTBY((D!B1:B3,D!C1:C3),{3;2;1}),1)', 'SORTBY', e('VALUE')],
  ['ARRAYTOTEXT(SORTBY(D!B1:B6,(D!C4:C6,D!C1:C3)),1)', 'SORTBY', s('{16;23;42;4;8;15}')],
  ['ARRAYTOTEXT(SORTBY(Jan:Mar!A1:A3,{9;8;7;6;5;4;3;2;1}),1)', 'SORTBY', s('{9;8;7;6;5;4;3;2;1}')],
  ['ARRAYTOTEXT(SORTBY(D!B1:B3,Jan:Mar!B1,-1),1)', 'SORTBY', s('{15;8;4}')],
  ['ROWS(SORTBY((D!B1:B3,D!C1:C3),{1;2;3;4;5;6}))', 'SORTBY', n(6)],
  ['ARRAYTOTEXT(SORTBY((D!B1:C1,D!B2:C2),{2;1}),1)', 'SORTBY', s('{8,7;4,2}')],
  ['ARRAYTOTEXT(SORTBY(D!B1:B3,D!C1:C3,1,(D!B1:B2,D!B3),-1),1)', 'SORTBY', s('{4;8;15}')],
  ['ARRAYTOTEXT(SORTBY(D!B1:B3,(D!C1,D!C2)),1)', 'SORTBY', e('VALUE')],
  ['ARRAYTOTEXT(SORTBY((D!B1:B2,Jan!A1),{2;1;3}),1)', 'SORTBY', e('VALUE')],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [f, fn] of CASES) add(fn, fn, f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(
  buildProbes().flatMap((p, i) => (CASES[i][2] ? [[p.id, CASES[i][2]]] : [])),
)
