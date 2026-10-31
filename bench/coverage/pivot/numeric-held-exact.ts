/** Repeats the held-out numeric inputs with binary64 transport, without the legacy Windows decimal parser. */
import { buildPivotCases as held } from './numeric-held.js'
export const buildPivotCases = () => held().filter(c => c.data === 'numeric-held')
