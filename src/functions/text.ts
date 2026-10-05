/**
 * Text functions.
 *
 * TEXT is this engine's own (`src/format.ts`).
 */
import type { FnDef } from '../args.js'
import { formatText } from '../format.js'
import { E, type CErr, type V } from '../values.js'

export const functions: Record<string, FnDef> = {
  TEXT: {
    args: 'v s',
    fn: (v: Exclude<V, CErr>, format: string) => formatText(v ?? 0, format) ?? E.VALUE,
  },
}
