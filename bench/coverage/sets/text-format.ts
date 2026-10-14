/**
 * Excel cases for TEXT with the format codes that the cell display needs too: exponents, elapsed time and
 * fractions of a second, colours and conditions (with their signs), padding `_x` and fills `*x`, fractions and
 * their signs, General among literals, currency and locale tags, four sections; in Excel en-US
 * (`test/excel/cases/text-format.json`). `text-format-ru` has the ru-RU codes.
 */
import type { Probe } from '../probes.js'
import { textFormatProbes } from './text-cases.js'

/** A formula's text literal. */

const NUMBERS = ['1234.567', '-1234.567', '0', '0.5', '-0.5', '-0.0000001', '1E+15', '"abc"']

const NUMBER_CODES = [
  '0.00E+00', '0.0E+0', '##0.0E+0', '0.00E-00', '0E+00', '#E+0', '0.00e+00', '0.00e-00', '[Red]0', '[Blue]0;[Red]-0', '[Color10]0.00', '[COLOR 10]0',
  '[Red][<0]0;0', '[>100]0.0;[<=-100]-0;0', '[<0]"neg";[>0]"pos";"zero"', '[=1]"one";"other"', '[>5]0', '[>100]"big"', '[<=9999999]###-####;(###) ###-####',
  '[<=0]0;0', '[<5]0;0', '[>-100]0;0', '[<>0]0;0', '0_)', '0_W', '_(0_)', '0*-', '*x0', '**0', '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', '# ?/?',
  '?/?', '0 ?/100', '# ??/16', 'General" pcs"', '"x"General', '[Red]General', '#,##0.00 [$₽-419]', '[$€-2] #,##0.00', '[$-409]0.00',
  '0;-0;0;"text: "@', '[Red]@', '0;;;', '0.0,', '#,##0,,"M"',
]

const DATES = ['45688.75', '0', '-1', '2958466', '45688.999999', '0.75/86400', '1.5']

const DATE_CODES = [
  '[h]:mm:ss', '[h]:mm', '[mm]:ss', '[ss]', '[hh]:mm', '[h]', 'h:mm:ss.0', 'mm:ss.00', 'mm:ss.000', 'h:mm a/p', '[Red]m/d/yyyy', '[$-419]mmmm', '[$-409]mmmm',
  'mmmmm', 'dddd', 'm/d/yyyy', 'yyyy-mm-dd hh:mm:ss',
]

export const buildProbes = (): Probe[] => textFormatProbes(NUMBER_CODES, NUMBERS, DATE_CODES, DATES)
