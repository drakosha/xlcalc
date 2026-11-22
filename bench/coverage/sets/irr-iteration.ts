import type { Probe, SheetCells } from '../probes.js'

const cashFlows = Array.from({ length: 61 }, (_, i) => (i < 14 ? -1 : 0.5))
const delayedInflow = Array.from({ length: 601 }, (_, i) => (i === 0 ? -100 : i === 600 ? 200 : 0))

export const SHEETS: SheetCells = {
  F: Object.fromEntries([
    ...cashFlows.map((value, i) => [`A${i + 1}`, value]),
    ...[...cashFlows].reverse().map((value, i) => [`B${i + 1}`, value]),
    ...delayedInflow.map((value, i) => [`C${i + 1}`, value]),
  ]),
}

export function buildProbes(): Probe[] {
  return [
    // One sign change after 14 outflows. The default guess needs a restart.
    'IRR(F!A1:A61)*12',
    'IRR(F!A1:A61,0.1)',
    'IRR(F!A1:A61,-0.9)',
    'IRR(F!A1:A61,100)',
    'IRR(F!B1:B61)',
    // An invalid discount divisor ends the calculation, even with a positive root.
    'IRR(F!C1:C601)',
    'IRR(F!C1:C601,0.05)',
    'IRR(F!C1:C601,5)',
    // Three roots. The guess controls which root the secant method reaches.
    'IRR({-100,320,-298,84})',
    'IRR({-100,320,-298,84},0.2)',
    'IRR({-100,320,-298,84},-0.5)',
    'IRR({-1,2.0001,-1},0)',
    'IRR({-1,2.0001,-1},0.01)',
    'IRR({-100,200,-100},0)',
    'IRR({-100,199,-100})',
    'IRR({0,0,0,-100,110},100)',
    'IRR({-100,0,110},100)',
  ].map((formula, i) => ({ id: `IRR#${i + 1}`, group: 'IRR', fn: 'IRR', formula }))
}
