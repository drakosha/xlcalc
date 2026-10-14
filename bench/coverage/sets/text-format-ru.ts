/**
 * Excel ru-RU cases for TEXT with the ru-RU format codes that `text-format` checks in en-US: exponents, elapsed
 * time and fractions of a second, colours (English and Russian names) and conditions, padding and fills,
 * fractions, General among literals, currency; in Excel ru-RU (`test/excel/cases/ru-RU/text-format-ru.json`).
 */
import type { Probe } from '../probes.js'
import { textFormatProbes } from './text-cases.js'

const NUMBERS = ['1234.567', '-1234.567', '0', '0.5', '-0.5', '1E+15', '"abc"']

const NUMBER_CODES = [
  '0,00E+00', '0,0E+0', '##0,0E+0', '0,00E-00', '0E+00', '[Red]0', '[Красный]0', '[Синий]0;[Красный]-0', '[Цвет10]0,00', '[Color10]0', '[Красный][<0]0;0',
  '[>100]0,0;[<=-100]-0;0', '[<0]"neg";[>0]"pos";"zero"', '0_)', '_(0_)', '0*-', '_(* # ##0_);_(* (# ##0);_(* "-"_);_(@_)', '# ?/?', '?/?',
  'Основной" шт"', '"x"Основной', '[Красный]Основной', '# ##0,00 [$₽-419]', '# ##0,00 ₽', '[$-409]0,00', '0;-0;0;"текст: "@',
]

const DATES = ['45688.75', '0', '-1', '45688.999999', '0.75/86400', '1.5']

const DATE_CODES = ['[ч]:мм:сс', '[ч]:мм', '[мм]:сс', '[сс]', '[ч]', 'ч:мм:сс,0', 'мм:сс,00', '[Красный]ДД.ММ.ГГГГ', '[$-409]ММММ', '[$-419]ММММ', 'ДД/ММ/ГГГГ', 'ДДДД']

export const buildProbes = (): Probe[] => textFormatProbes(NUMBER_CODES, NUMBERS, DATE_CODES, DATES)
