/**
 * Held out, predicted by the engine before Excel ran it: the number a percent with decimals is stored as. The engine
 * moves the decimal point of the typed text (`1.1%` is 0.011, as `Number("1.1e-2")`), where dividing the number read by
 * 100 gives 0.011000000000000001; the texts are ones where the two part, in both locales' forms, with a sign,
 * parentheses, a leading `%`, groups, an exponent and money. Fractions (`1 2/3`, with percent too) are where a whole part
 * plus the fraction and the fraction over its denominator part. Plain decimals and the scientific form are the control.
 */
import { type EntryCase, entryCases } from './cases.js'

const GENERAL: Record<string, string[]> = {
  pen: ['1.1%', '0.7%', '1.255%', '99.9%', '8.2%', '33.3%', '57.1%', '0.035%', '1.005%', '0.07%', '123.45%', '-0.7%', '(1.1%)', '%1.1', '1,000.7%', '1.1E+0%', '$1.1%'],
  pru: ['1,1%', '0,7%', '1,255%', '99,9%', '8,2%', '33,3%', '57,1%', '0,035%', '1,005%', '0,07%', '123,45%', '-0,7%', '(1,1%)', '%1,1', '1 000,7%', '1,1E+0%', '1,1 ₽%'],
  frac: ['1 2/3', '1 5/6', '2 4/7', '1 4/11', '1 2/3%', '1 1/3%', '0 1/3%'],
  plain: ['0.011', '0,011', '1.1E-2', '1,1E-2', '0.007', '0,007'],
}

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL)
