/**
 * Open cases of the random check (the `fuzz` set): differences from Excel whose cause lies in engine files
 * outside the case work (the compiler, the runtime, the array layer), each with one line why and the engine's
 * answer pinned, as KNOWN does. The lead closes them; an entry whose case matches Excel fails the tests.
 */
import type { Val } from '../bench/coverage/formulajs.js'

export const OPEN: Readonly<Record<string, { why: string; engine: Val }>> = {}
