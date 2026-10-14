/**
 * Excel ru-RU cases for TEXT: colour names, month names after a day, the system long date (`test/excel/cases/ru-RU/text-format-ru-more.json`).
 */
import type { Probe } from '../probes.js'
import { textProbes } from './text-cases.js'

const CASES: [value: string, format: string][] = [
  ['5', '[Черный]0'], ['5', '[Голубой]0'], ['5', '[Зеленый]0'], ['5', '[Фиолетовый]0'], ['5', '[Белый]0'], ['5', '[Желтый]0'], ['5', '[Цвет 10]0'], ['5', '[ЦВЕТ10]0'],
  ['5', '[красный]0'], ['5', '[Зелёный]0'], ['45688', 'Д ММММ ГГГГ'], ['45688', 'ДД ММММ'], ['45688', '[$-F800]ДДДД, ММММ ДД, ГГГГ'], ['45688.75', '[$-F400]ч:мм:сс'],
  ['1000000000', '?/?'], ['1/3', '@'],
]

export const buildProbes = (): Probe[] => textProbes(CASES)
