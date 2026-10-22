/**
 * The values r2/values changed without a recorded answer, held out in Excel ru-RU alone (`probes values-ru-held ru`):
 * FIXED and DOLLAR from 1E+21 on and digit formats past the 15th significant digit with the locale's separators, a
 * colour name in brackets beside a point before a digit placeholder (`[Красный]0.00`, refused as `0.00` is; the
 * colours whose names hold a date letter and two without), and a text section ending a time code (`ч:мм;@`). Answers:
 * `test/excel/cases/ru-RU/values-ru-held.json`; the engine's answers, `values-ru-held.ru.predicted.json`, were committed
 * before Excel ran the set.
 */
import { heldCases } from './arrays/common.js'

export const { buildProbes } = heldCases([
  ['FIXED', 'FIXED(1E+21)'],
  ['FIXED', 'FIXED(1E+21,1,TRUE)'],
  ['FIXED', 'FIXED(2^60,0)'],
  ['FIXED', 'FIXED(1/3,20)'],
  ['DOLLAR', 'DOLLAR(1E+21)'],
  ['DOLLAR', 'DOLLAR(-1E+21,0)'],
  ['TEXT', 'TEXT(1/3,"0,00000000000000000000")'],
  ['TEXT', 'TEXT(2^60,"# ##0")'],
  ['TEXT', 'TEXT(793853.646,"0,0000000000")'],
  ['TEXT', 'TEXT(1234.5,"0.00")'],
  ['TEXT', 'TEXT(1234.5,"[Красный]0.00")'],
  ['TEXT', 'TEXT(1234.5,"[Синий]0.00")'],
  ['TEXT', 'TEXT(1234.5,"[Черный]0.00")'],
  ['TEXT', 'TEXT(1234.5,"[Голубой]0.00")'],
  ['TEXT', 'TEXT(1234.5,"[Зеленый]0.00")'],
  ['TEXT', 'TEXT(1234.5,"[Желтый]0.00")'],
  ['TEXT', 'TEXT(1234.5,"[Красный]#.##")'],
  ['TEXT', 'TEXT(1234.5,"[Красный]0,00")'],
  ['TEXT', 'TEXT(45688,"[Красный]ДД.ММ.ГГГГ")'],
  ['TEXT', 'TEXT(1.25,"[Черный][ч]:мм")'],
  ['TEXT', 'TEXT(-0.25,"ч:мм;@")'],
  ['TEXT', 'TEXT(-5,"0;@")'],
])
