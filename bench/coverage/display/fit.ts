/**
 * What does not fit a column, and what a format does with the width: booleans, errors and text in narrow columns,
 * formatted numbers and dates of mixed characters near the edge, fills `*x` and padding `_x` against the width,
 * shown text longer than 255 characters, merged areas, the default column, colours of every kind of section,
 * date separators under locale tags, sections chosen by conditions that fall back to General.
 */
import { collector } from './cases.js'

const px = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => (from + i) / 6)

export function buildDisplayCases() {
  const { out, add } = collector()
  for (const width of px(3, 50)) for (const v of [true, false, '=NA()', '=1/0', '=SQRT(-1)', '="a"+1']) add('word', v, undefined, { width })
  for (const width of px(3, 12)) add('text', 'abcdefgh', undefined, { width })
  for (const width of px(60, 75)) add('mixed', 1234567.891, '#,##0.00', { width })
  for (const width of px(60, 75)) add('mixed', -1234.5, '$#,##0.00_);($#,##0.00)', { width })
  for (const width of px(90, 110)) add('mixed', 45688, 'mmmm d, yyyy', { width })
  for (const width of px(40, 60)) add('mixed', 45688.75, 'h:mm AM/PM', { width })
  for (const width of px(25, 45)) add('mixed', 0.1234, '0.0%', { width })
  for (const width of px(30, 45)) add('mixed', 1234.567, '0.00E+00', { width })
  for (const width of px(30, 45)) add('pad', 1235, '0_W', { width })
  for (const width of px(20, 40)) add('fill', 5, '0*-', { width })
  for (const width of px(20, 40)) add('fill', 'ab', '@*.', { width })
  for (const width of px(20, 40)) add('fill', 1234, '* #,##0', { width })
  for (const width of px(20, 40)) add('fill', -12, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', { width })
  for (const width of px(20, 40)) add('fill', 7, '*x0*y', { width })
  for (const width of [80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91]) add('long', 5, '0* ', { width })
  add('long', 1e300, '0', { width: 255 })
  add('long', 1e250, '0', { width: 255 })
  for (const span of [2, 3]) for (const width of px(3, 12)) add('merged', 1e20, '0', { width, span })
  for (const span of [2, 3]) for (const width of px(6, 12)) add('merged', 123456789, undefined, { width, span })
  for (const v of [123456789, 1234567.891, 12345678, 0.000123456, -1234567, 'a long piece of text', true, '=NA()']) add('default', v)
  for (const [v, f] of [
    ['abc', '0;0;0;[Red]@'], ['abc', '[Red]@'], ['abc', '[Blue]General'], [5, '[Red]0;[Blue]0;[Green]0;[Magenta]@'], [-5, '[Red]0;[Blue]0;[Green]0;[Magenta]@'],
    [0, '[Red]0;[Blue]0;[Green]0;[Magenta]@'], ['x', '[Red]0;[Blue]0;[Green]0;[Magenta]@'], [0, '[Red]0'], [0, '[Red]0;[Blue]0'], [0, '[Red]0;[Blue]-0;[Green]0'],
    [true, '[Red]0'], ['=NA()', '[Red]0'], [-0.0000001, '[Red]0;[Blue]0'], [0, '[Red][=0]0;[Blue]0'], [0, '[Red][<=0]0;[Blue]0'], [5, '[Red][>5]0;[Blue][<0]0;[Green]0'],
  ] as const) add('color', v, f, { width: 20 })
  for (const [v, f] of [
    [45688, '[$-409]m/d/yyyy'], [45688, '[$-409]dd/mm/yyyy'], [45688, 'dd-mm-yyyy'], [45688, 'd/m'], [45688, '[$-419]m/d/yyyy'], [45688.75, 'h:mm:ss'], [45688, 'mmm/yy'],
    [45688, '[$-409]mmmm'], [45688, 'dd\\/mm\\/yyyy'], [45688, 'dd"/"mm"/"yyyy'], [1.5, '0/0'], [45688, '[$-F800]dddd, mmmm dd, yyyy'],
  ] as const) add('separator', v, f, { width: 30 })
  for (const width of px(20, 70)) add('fallback', -1234.567, '[>5]0', { width })
  for (const [v, f] of [
    [2958465.999999, 'm/d/yyyy'], [2958465.999999, 'm/d/yyyy h:mm:ss.000'], [-0.0000001, 'm/d/yyyy'], [-0.0000001, 'h:mm'], [2958465.5, '[h]:mm'], [-1, '[ss]'], [-1, '[h]'], [-1, '[mm]'],
    [1e15, '?/?'], [2147483647, '?/?'], [2147483648, '?/?'], [1e10 + 0.5, '?/?'], [99999999999, '# ?/?'],
  ] as const) add('edge', v, f, { width: 40 })
  return out
}
