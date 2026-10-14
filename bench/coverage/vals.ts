/** Excel's answers written in the case sets: a number, a text, a logical value, an error. */
import type { Val } from './formulajs.js'

export const n = (v: number): Val => ({ t: 'n', v })
export const s = (v: string): Val => ({ t: 's', v })
export const b = (v: boolean): Val => ({ t: 'b', v })
export const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })
