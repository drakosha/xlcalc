import type { FilterCriteria } from '../src/contract.js'

/** The legacy filter-mode marker, now represented as criteria when read back. */
export const ALL_ROWS: FilterCriteria = { columns: [{ column: 0, filter: { type: 'custom', conditions: [{ value: '' }, { operator: 'notEqual', value: '' }] } }] }
