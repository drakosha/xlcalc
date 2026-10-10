/**
 * AutoFit of text, character by character: each character of the width table (`display/widths.ts`) alone and repeated
 * to about 600 points, where Excel measures text wider than it shows it; Russian words, among them the texts of
 * booleans and errors in ru-RU, as text.
 */
import { glyph } from '../../../src/numfmt/glyphs.js'
import { GLYPHS } from '../display/widths.js'
import { collector, type FitCase } from './cases.js'

const WORDS = ['ИСТИНА', 'ЛОЖЬ', '#Н/Д', '#ЗНАЧ!', '#ИМЯ?', '#ДЕЛ/0!', '#ЧИСЛО!', '#ССЫЛКА!', '#ПУСТО!', 'TRUE', '#N/A', 'Иван', 'Наименование', 'Сумма', 'Итого', 'Дата', 'ДДДД', 'ИИИИ', 'НННН', 'АААА']

export function buildFitCases(): FitCase[] {
  const { out, add } = collector()
  for (const g of GLYPHS) add('one', { cells: { 1: { value: g } } })
  for (const g of GLYPHS) add('many', { cells: { 1: { value: g.repeat(Math.floor(600 / glyph(g))) } } })
  for (const w of WORDS) add('word', { cells: { 1: { value: w } } })
  return out
}
