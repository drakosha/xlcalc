/**
 * Percent text read as a number, held out: the engine moves the decimal point of the text (`"1.1%"` is 0.011, as
 * `Number("1.1e-2")`), where dividing the number read by 100 gives 0.011000000000000001. The difference from the decimal
 * literal is magnified by 1E+20 (the last operation a product, so Excel's rounding of a difference near 0 does not
 * hide it): 0 when the text reads as the literal. In VALUE, operators and a numeric argument (ABS), with each locale's
 * decimal separator and with none (`11E-1%`); a whole part with a fraction (`1 2/3`, with percent too); NUMBERVALUE and
 * the percent operator, which the engine keeps dividing by 100. Computed in Excel en-US
 * (`test/excel/cases/pct-text.json`) and ru-RU (`test/excel/cases/ru-RU/pct-text.json`); `PREDICTED` and
 * `pct-text.ru.predicted.json`, the engine's answers, were committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

/** Percent text and the decimal literal it reads as: in en-US's form, ru-RU's, and with no decimal separator. */
const TEXTS: [string, string][] = [
  ['1.1%', '0.011'],
  ['0.7%', '0.007'],
  ['1.255%', '0.01255'],
  ['99.9%', '0.999'],
  ['0.035%', '0.00035'],
  ['1.005%', '0.01005'],
  ['123.45%', '1.2345'],
  ['(1.1%)', '-0.011'],
  ['%0.7', '0.007'],
  ['$1.1%', '0.011'],
  ['1,1%', '0.011'],
  ['0,7%', '0.007'],
  ['1,255%', '0.01255'],
  ['99,9%', '0.999'],
  ['123,45%', '1.2345'],
  ['(1,1%)', '-0.011'],
  ['%0,7', '0.007'],
  ['11E-1%', '0.011'],
  ['7E-1%', '0.007'],
  ['1 2/3', '1+2/3'],
  ['2 4/7', '2+4/7'],
  ['1 2/3%', '(1+2/3)/100'],
]

const OPERATOR = ['(1.1%-0.011)*1E+20', '(0.7%-0.007)*1E+20', '(99.9%-0.999)*1E+20', '((11/10)%-0.011)*1E+20', '(9%%-0.0009)*1E+20']

const NUMBERVALUE = ['(NUMBERVALUE("1.1%",".",",")-0.011)*1E+20', '(NUMBERVALUE("1,1%",",",".")-0.011)*1E+20', '(NUMBERVALUE("0.7%",".")-0.007)*1E+20', '(NUMBERVALUE("9%%",".")-0.0009)*1E+20', '(NUMBERVALUE("1.1%%",".")-0.00011)*1E+20']

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [t, n] of TEXTS) {
    add('VALUE', 'VALUE', `(VALUE("${t}")-(${n}))*1E+20`)
    add('operators', 'operators', `("${t}"*1-(${n}))*1E+20`)
    add('ABS', 'ABS', `(ABS("${t}")-ABS(${n}))*1E+20`)
  }
  add('percent', 'operators', ...OPERATOR)
  add('NUMBERVALUE', 'NUMBERVALUE', ...NUMBERVALUE)
  return out
}

const n = (v: number): Val => ({ t: 'n', v })
const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })

/** The engine's answers in en-US, predicted before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = {
  'VALUE#1': n(0),
  'operators#1': n(0),
  'ABS#1': n(0),
  'VALUE#2': n(0),
  'operators#2': n(0),
  'ABS#2': n(0),
  'VALUE#3': n(0),
  'operators#3': n(0),
  'ABS#3': n(0),
  'VALUE#4': n(0),
  'operators#4': n(0),
  'ABS#4': n(0),
  'VALUE#5': n(0),
  'operators#5': n(0),
  'ABS#5': n(0),
  'VALUE#6': n(0),
  'operators#6': n(0),
  'ABS#6': n(0),
  'VALUE#7': n(0),
  'operators#7': n(0),
  'ABS#7': n(0),
  'VALUE#8': n(0),
  'operators#8': n(0),
  'ABS#8': n(0),
  'VALUE#9': n(0),
  'operators#9': n(0),
  'ABS#9': n(0),
  'VALUE#10': n(0),
  'operators#10': n(0),
  'ABS#10': n(0),
  'VALUE#11': e('VALUE'),
  'operators#11': e('VALUE'),
  'ABS#11': e('VALUE'),
  'VALUE#12': e('VALUE'),
  'operators#12': e('VALUE'),
  'ABS#12': e('VALUE'),
  'VALUE#13': n(1.253745e+21),
  'operators#13': n(1.253745e+21),
  'ABS#13': n(1.253745e+21),
  'VALUE#14': e('VALUE'),
  'operators#14': e('VALUE'),
  'ABS#14': e('VALUE'),
  'VALUE#15': e('VALUE'),
  'operators#15': e('VALUE'),
  'ABS#15': e('VALUE'),
  'VALUE#16': e('VALUE'),
  'operators#16': e('VALUE'),
  'ABS#16': e('VALUE'),
  'VALUE#17': e('VALUE'),
  'operators#17': e('VALUE'),
  'ABS#17': e('VALUE'),
  'VALUE#18': n(0),
  'operators#18': n(0),
  'ABS#18': n(0),
  'VALUE#19': n(0),
  'operators#19': n(0),
  'ABS#19': n(0),
  'VALUE#20': n(0),
  'operators#20': n(0),
  'ABS#20': n(0),
  'VALUE#21': n(0),
  'operators#21': n(0),
  'ABS#21': n(0),
  'VALUE#22': n(0),
  'operators#22': n(0),
  'ABS#22': n(0),
  'percent#1': n(173.472347597681),
  'percent#2': n(-86.7361737988404),
  'percent#3': n(11102.2302462516),
  'percent#4': n(173.472347597681),
  'percent#5': n(0),
  'NUMBERVALUE#1': n(173.472347597681),
  'NUMBERVALUE#2': n(173.472347597681),
  'NUMBERVALUE#3': n(-86.7361737988404),
  'NUMBERVALUE#4': n(0),
  'NUMBERVALUE#5': n(0),
}
