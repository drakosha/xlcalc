/** Helpers shared by the text functions. */
import { E, type V } from '../../values.js'

/** Excel's limit on the length of a text value. */
export const MAX_TEXT = 32_767

export const capped = (s: string): V => (s.length > MAX_TEXT ? E.VALUE : s)
