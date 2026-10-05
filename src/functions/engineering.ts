/**
 * Engineering functions. None is in the whitelist yet: formula.js's engineering functions need the argument
 * layer's rules (no booleans: kind `x`; upper-case digits and complex text: `upperHex`, `complexText` in
 * `src/args.ts`) and domain checks, verified against the Excel cases before they are added here.
 */
import type { FnDef } from '../args.js'

export const functions: Record<string, FnDef> = {}
