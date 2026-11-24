import type { BookInput, CellStyle, WorkbookTableStyles } from './contract.js'
import type { Palette } from './xlsx/styles/color.js'
import type { Rect } from './structure/area.js'
import type { Dxf } from './xlsx/styles/dxf.js'

export interface TableStyleFiles {
  read(xml: string | undefined, colors: Palette, dxfs: Dxf[]): WorkbookTableStyles | undefined
  check(data: unknown, path: string): void
  write(data: WorkbookTableStyles, dxf: (xml: string) => number): { styles?: string; theme?: string }
}

export let tableStyleFiles: TableStyleFiles | undefined

export function setTableStyleFiles(files: TableStyleFiles): void {
  tableStyleFiles = files
}

export interface TableStyleWindow {
  at(row: number, col: number): CellStyle | undefined
  resolve(base: CellStyle, direct: CellStyle | undefined, conditional: CellStyle | undefined): number
  readonly styles: Record<number, CellStyle> | undefined
}

export type OpenTableStyleWindow = (book: BookInput, sheet: number, boxes: Rect[]) => TableStyleWindow | undefined
export let openTableStyleWindow: OpenTableStyleWindow | undefined

const listeners = new Set<() => void>()

export function setTableStyleWindow(open: OpenTableStyleWindow): void {
  if (openTableStyleWindow === open) return
  openTableStyleWindow = open
  for (const listener of listeners) listener()
}

export function onTableStylesLoaded(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}
