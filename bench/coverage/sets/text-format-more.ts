/**
 * Excel en-US cases for TEXT: improper fractions of large numbers, a number under `@`, long text from a number
 * format, a value that no conditional section takes (`test/excel/cases/text-format-more.json`).
 */
import type { Probe } from '../probes.js'
import { textProbes } from './text-cases.js'

const CASES: [value: string, format: string][] = [
  ['99999999', '?/?'], ['100000000', '?/?'], ['999999999', '?/?'], ['1000000000', '?/?'], ['1000000001', '?/?'], ['2000000000', '?/?'], ['123456789.5', '?/?'],
  ['1000000000.5', '?/?'], ['1000000000.5', '??/??'], ['1/3', '@'], ['1234567.891', '@'], ['1E+300', '0'], ['1E+250', '0'], ['5', '0*-'], ['45688', '[$-F800]dddd, mmmm dd, yyyy'],
  ['45688.75', '[$-F400]h:mm:ss AM/PM'], ['45688', 'd mmmm yyyy'], ['0', '[<0]0;[>0]0'], ['0', '[>5]0;[<-5]0'], ['3', '[>5]0;[<-5]0;'], ['3', '[>5]0;[<-5]0;0'],
  ['-1', '[>5]0'],
]

export const buildProbes = (): Probe[] => textProbes(CASES)
