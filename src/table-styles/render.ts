import type { BookInput, BorderLine, CellStyle, WorkbookTableStyles } from '../contract.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { keyOf } from '../style-table.js'
import { anyMeets, hasCell, type Rect } from '../structure/area.js'
import { tableKey } from '../tables/model.js'
import { DEFAULT_TABLE_STYLE } from '../tables/check.js'
import type { TableStyleWindow } from '../table-style-hook.js'
import { colorOf, type Palette } from '../xlsx/styles/color.js'
import { dxfCellStyle, readDxf } from '../xlsx/styles/dxf.js'
import { checkTableStyleData } from '../input/table-style-data.js'
import { LINES } from '../input/styles.js'
import { attributes, element } from '../xlsx/xml/scan.js'
import { PRESET_DIFFERENTIAL_FORMATS, PRESET_TABLE_STYLES } from './presets.js'
import { pivotArea, tableArea, type StyledArea } from './regions.js'

interface Differential {
  style: CellStyle
  horizontal?: BorderLine
  vertical?: BorderLine
}

interface ElementStyle extends Differential {
  size?: number
}

type Definition = ReadonlyMap<string, ElementStyle>
type Side = keyof NonNullable<CellStyle['border']>
const PRIORITY = new Map([
  'wholeTable', 'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe',
  'blankRow', 'firstSubtotalRow', 'secondSubtotalRow', 'thirdSubtotalRow',
  'firstSubtotalColumn', 'secondSubtotalColumn', 'thirdSubtotalColumn',
  'lastColumn', 'firstColumn', 'firstRowSubheading', 'secondRowSubheading', 'thirdRowSubheading',
  'totalRow', 'lastTotalCell', 'firstTotalCell', 'headerRow',
  'firstColumnSubheading', 'secondColumnSubheading', 'thirdColumnSubheading',
  'lastHeaderCell', 'firstHeaderCell', 'pageFieldLabels', 'pageFieldValues',
].map((type, index) => [type, index]))
const SIDES: readonly Side[] = ['top', 'right', 'bottom', 'left']
const NEIGHBORS = [
  ['top', 'bottom', -1, 0],
  ['right', 'left', 0, 1],
  ['bottom', 'top', 1, 0],
  ['left', 'right', 0, -1],
] as const

// The palette Excel for Windows uses for a workbook without a theme part.
const OFFICE = ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E', '467886', '96607D']
const presetNames = new Map(Object.keys(PRESET_TABLE_STYLES).map((name) => [tableKey(name), name]))
const presetCache = new WeakMap<string[], { formats: Map<number, Differential>; styles: Map<string, Definition> }>()

function presetFormat(xml: string, palette: Palette): Differential {
  const dxf = element(xml, 'dxf')!
  const parsed = readDxf(dxf, palette)
  const style = dxfCellStyle(parsed)
  if (!style.fill && parsed.pattern) style.fill = { color: parsed.pattern }
  const border = element(dxf.body ?? '', 'border')?.body ?? ''
  const inner = (side: string): BorderLine | undefined => {
    const line = element(border, side)
    const kind = attributes(line?.attrs ?? '').style as BorderLine['style']
    if (!LINES.includes(kind)) return
    const color = colorOf(element(line?.body ?? '', 'color'), palette)
    return color ? { style: kind, color } : { style: kind }
  }
  return { style, horizontal: inner('horizontal'), vertical: inner('vertical') }
}

function preset(name: string, theme: string[]): Definition | undefined {
  const spelling = presetNames.get(name)
  if (!spelling) return
  let cache = presetCache.get(theme)
  if (!cache) presetCache.set(theme, (cache = { formats: new Map(), styles: new Map() }))
  let definition = cache.styles.get(spelling)
  if (!definition) {
    const palette = { theme: theme === OFFICE ? OFFICE : theme.map((c) => c.slice(1)), indexed: [] }
    definition = new Map(PRESET_TABLE_STYLES[spelling].map(([type, size, index]) => {
      let format = cache.formats.get(index)
      if (!format) cache.formats.set(index, (format = presetFormat(PRESET_DIFFERENTIAL_FORMATS[index], palette)))
      return [type, { ...format, size }]
    }))
    cache.styles.set(spelling, definition)
  }
  return definition
}

/** Style fields override independently, including individual font attributes and border sides. */
function overlay(base: CellStyle, top: CellStyle | undefined): CellStyle {
  if (!top) return base
  return {
    ...(base.fill || top.fill ? { fill: top.fill ?? base.fill } : {}),
    ...(base.font || top.font ? { font: { ...base.font, ...top.font } } : {}),
    ...(base.border || top.border ? { border: { ...base.border, ...top.border } } : {}),
    ...(base.alignment || top.alignment ? { alignment: { ...base.alignment, ...top.alignment } } : {}),
    ...(base.protection || top.protection ? { protection: { ...base.protection, ...top.protection } } : {}),
  }
}

interface Area {
  area: StyledArea
  definition: Definition
}
interface Drawn {
  area: Area
  style: CellStyle
  rank: Partial<Record<Side, number>>
}

/** A region's outer edges run round its box. Its inner edges are between cells of that box. */
function draw({ area, definition }: Area, r: number, c: number): Omit<Drawn, 'area'> {
  const rank: Drawn['rank'] = {}
  let style: CellStyle = {}
  const regions = area.regions(r, c, (name) => definition.get(name)?.size ?? 1)
  regions.forEach(({ type, box }) => {
    const format = definition.get(type)
    if (!format) return
    const { border, ...look } = format.style
    const sides = [
      r === box.r1 ? border?.top : format.horizontal,
      c === box.c2 ? border?.right : format.vertical,
      r === box.r2 ? border?.bottom : format.horizontal,
      c === box.c1 ? border?.left : format.vertical,
    ]
    const edges: CellStyle['border'] = {}
    SIDES.forEach((side, index) => {
      const line = sides[index]
      if (line) {
        edges[side] = line
        rank[side] = PRIORITY.get(type)
      }
    })
    style = overlay(style, Object.keys(edges).length ? { ...look, border: edges } : look)
  })
  return { style, rank }
}

function definitions(data: WorkbookTableStyles | undefined) {
  const custom = new Map(Object.entries(data?.custom ?? {}).map(([name, def]) => [tableKey(name), new Map(def.elements.map((e) => [e.type, e]))]))
  const theme = data?.theme ?? OFFICE
  return (name: string) => {
    const key = tableKey(name)
    return custom.get(key) ?? preset(key, theme)
  }
}

export function tableStyleWindow(book: BookInput, sheetIndex: number, rects: Rect[]): TableStyleWindow | undefined {
  const sheet = book.sheets[sheetIndex]
  if (!sheet.tables?.length && !sheet.pivotTables?.length) return
  if (book.tableStyles !== undefined) checkTableStyleData(book.tableStyles, 'book.tableStyles')
  const definitionOf = definitions(book.tableStyles)
  const areas: Area[] = []
  // A border is visible from either cell sharing the edge, even just outside a styled area.
  const nearby = rects.map(({ r1, r2, c1, c2 }) => ({ r1: r1 - 1, r2: r2 + 1, c1: c1 - 1, c2: c2 + 1 }))
  const add = (name: string | undefined, locate: () => StyledArea) => {
    const definition = name ? definitionOf(name) : undefined
    if (!definition?.size) return
    const area = locate()
    if (anyMeets(area.bounds, nearby)) areas.push({ area, definition })
  }
  for (const table of sheet.tables ?? []) add((table.style ?? DEFAULT_TABLE_STYLE).name, () => tableArea(table))
  for (const table of sheet.pivotTables ?? []) add(table.style?.name, () => pivotArea(table))
  if (!areas.length) return

  const cells = new Map<number, Drawn | null>()
  const drawn = (r: number, c: number): Drawn | null => {
    if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return null
    const key = r * MAX_COLS + c
    const cached = cells.get(key)
    if (cached !== undefined) return cached
    const area = areas.find(({ area }) => area.bounds.some((bounds) => hasCell(bounds, r, c)))
    const look = area ? { area, ...draw(area, r, c) } : null
    cells.set(key, look)
    return look
  }
  const ids = new Map<string, number>()
  const styles: Record<number, CellStyle> = {}
  return {
    get styles() {
      return ids.size ? styles : undefined
    },
    at(r, c) {
      const here = drawn(r, c)
      const own = here?.style
      let border = own?.border
      for (const [side, otherSide, dr, dc] of NEIGHBORS) {
        const next = drawn(r + dr, c + dc)
        const nextSide = next?.style.border?.[otherSide]
        if (next && nextSide && (!here || next.area === here.area && (next.rank[otherSide] ?? -1) > (here.rank[side] ?? -1))) {
          border = { ...border, [side]: nextSide }
        }
      }
      if (border !== own?.border) return { ...own, border }
      return own && Object.keys(own).length ? own : undefined
    },
    resolve(base, direct, conditional) {
      const style = overlay(overlay(base, direct), conditional)
      const key = keyOf(style)
      let id = ids.get(key)
      if (id === undefined) {
        id = ids.size
        ids.set(key, id)
        styles[id] = JSON.parse(JSON.stringify(style)) as CellStyle
      }
      return id
    },
  }
}
