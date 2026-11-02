/**
 * The built-in formats Excel shows by the regional settings (money, dates and times, negatives in parentheses), and the
 * East Asian dates 50 to 58; the others every locale shows alike, so a file stores their codes by number (a locale may
 * still write one otherwise, as ru-RU quotes a fraction's space). A module apart from the tables of codes
 * (`builtin.ts`), which `xlcalc/text` takes without it.
 */
const span = (lo: number, hi: number) => Array.from({ length: hi - lo + 1 }, (_, i) => lo + i)

export const REGIONAL: ReadonlySet<number> = new Set([...span(5, 8), ...span(14, 44), ...span(50, 58)])
