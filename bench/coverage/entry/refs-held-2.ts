/**
 * Held out after MOD was found to pass a format on (the set `refs-held`), predicted by the engine before Excel ran it:
 * MOD over money, a date and a percent second, and functions already measured over other formats.
 */
import { type EntryCase, entryCases } from './cases.js'

export { data } from './refs-held.js'

export const buildEntryCases = (): EntryCase[] => entryCases({ mod: ['=MOD(D3,7)', '=MOD(D6,D12)', '=MOD(D11,7)', '=ABS(D11)', '=TRUNC(D3)', '=ROUNDDOWN(D11,0)', '=MOD(D5,1)-D6'] })
