import type { FitCase } from './cases.js'

export const CELL_FONTS = ['Arial', 'Arial Cyr', 'Times New Roman'].flatMap((name) =>
  Array.from({ length: 7 }, (_, index) => {
    const size = index + 8
    return { id: `${name.toLowerCase().replaceAll(' ', '-')}-${size}`, font: { name, size } }
  }),
)

export function buildFitCases(): FitCase[] {
  return CELL_FONTS.flatMap(({ id, font }) => [
    { id: `${id}-latin`, cells: { 1: { value: 'MWilWAV 0123456789'.repeat(2), look: { font } } } },
    { id: `${id}-cyrillic`, cells: { 1: { value: 'Итог за ноябрь, 123 ₽'.repeat(2), look: { font } } } },
    { id: `${id}-number`, readWidth: true, column: { width: 3.8 }, cells: { 1: { value: 8.25, format: '0.00', look: { font } } } },
  ])
}
