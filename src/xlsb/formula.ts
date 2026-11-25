import { quoteSheetIfNeeded } from '../a1.js'
import { wordKey } from '../collate.js'
import { generalText } from '../numfmt/general.js'
import { quoteSheet } from '../sheet-pattern.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { linkedTableText, type TableSpec } from '../syntax/table.js'
import { XLFN, XLWS } from '../xlsx/formulas/prefixes.js'
import { address, errorText } from './fields.js'
import { FTAB } from './ftab.js'
import type { Context } from './parts.js'
import type { Payload } from './records.js'

export interface Formula {
  tokens: Payload
  extra: Payload
  xls?: XlsFormula
}

export interface XlsFormula {
  old?: 2 | 3 | 4 | 5
  links?: readonly LegacyLink[]
  string(payload: Payload): string
  array(payload: Payload): string | undefined
  label?(type: number, tokens: Payload, extra: Payload, readArea: () => string): string | undefined
}

export interface LegacyLink {
  number?: number
  sheet?: string
  current?: true
  names?: { name: string; sheet?: string }[]
}

export interface Defined {
  name: string
  scope: number
  hidden: boolean
  empty: boolean
  future: boolean
  ref: Formula
}

type Operand = string | { text: string; name: Defined }
const render = (operand: Operand): string => typeof operand === 'string' ? operand : operand.text
const nativeFunctions = new Set(Object.values(FTAB).map((entry) => entry?.[0]))

function functionName(operand: Operand): string {
  if (typeof operand === 'string' || !operand.name.empty) return render(operand)
  // Excel also marks dotted and dotless I spellings of worksheet functions as _xludf.
  const name = wordKey(operand.name.name).replace(/[ıİ]/g, 'I')
  if (!nativeFunctions.has(name) && !XLFN.has(name) && !XLWS.has(name)) return operand.text
  return operand.text.replace(/^\s*/, '$&_xludf.')
}

export type FormulaText = { text: string } | { anchor: string } | { table: string } | undefined

export function formula(payload: Payload): Formula {
  const tokens = payload.section(payload.u32())
  const extra = payload.section(payload.u32())
  return { tokens, extra }
}

const BINARY = ['+', '-', '*', '/', '^', '&', '<', '<=', '=', '>=', '>', '<>', ' ', ',', ':']
export const textLiteral = (value: string) => `"${value.replace(/"/g, '""')}"`
const AREAS: Readonly<Partial<Record<number, TableSpec['areas']>>> = {
  0: [], 1: ['all'], 2: ['headers'], 4: ['data'], 6: ['headers', 'data'], 8: ['totals'], 12: ['data', 'totals'], 16: ['row'],
}

/** The signed offset and the two relativity bits are separate for each coordinate. */
function coordinate(value: number, size: number, relative: boolean, origin?: number): number {
  const index = value & (size - 1)
  return relative && origin !== undefined ? (origin + index) & (size - 1) : index
}

function ref(payload: Payload, rowBits: number, colBits: number, xls: XlsFormula | undefined, row?: number, col?: number) {
  const flags = xls?.old ? rowBits : colBits
  const relC = Boolean(flags & 0x4000)
  const relR = Boolean(flags & 0x8000)
  const r = coordinate(rowBits, xls?.old ? 16384 : xls ? 65536 : MAX_ROWS, relR, row)
  const c = coordinate(colBits, xls ? 256 : MAX_COLS, relC, col)
  const a = address(payload, r, c)
  const rowText = String(r + 1)
  const letters = a.slice(0, -rowText.length)
  return { r, c, row: (relR ? '' : '$') + rowText, col: (relC ? '' : '$') + letters }
}

function area(payload: Payload, xls: XlsFormula | undefined, linked: boolean, row?: number, col?: number): string {
  const r1 = xls ? payload.u16() : payload.u32()
  const r2 = xls ? payload.u16() : payload.u32()
  const c1 = xls?.old ? payload.u8() : payload.u16()
  const c2 = xls?.old ? payload.u8() : payload.u16()
  const a = ref(payload, r1, c1, xls, row, col)
  const b = ref(payload, r2, c2, xls, row, col)
  const lastRow = xls?.old ? 16383 : xls ? 65535 : MAX_ROWS - 1
  const fullHeight = a.r === 0 && b.r === lastRow
  const fullWidth = a.c === 0 && b.c === (xls ? 255 : MAX_COLS - 1)
  if (xls && linked) {
    // A linked workbook keeps the legacy grid bounds when Excel upgrades the referencing workbook.
    if (fullHeight && fullWidth) return `${a.col}:${b.col}`
    const end = fullHeight && xls.old === 5 ? `${b.row.startsWith('$') ? '$' : ''}65536` : b.row
    return `${a.col}${a.row}:${b.col}${end}`
  }
  // Excel uses whole rows when the range spans both axes of its own sheet.
  if (fullWidth) {
    // Excel expands a reference over the entire binary sheet when it opens the workbook.
    const end = xls && fullHeight ? `${b.row.startsWith('$') ? '$' : ''}${MAX_ROWS}` : b.row
    return `${a.row}:${end}`
  }
  if (fullHeight) return `${a.col}:${b.col}`
  return `${a.col}${a.row}:${b.col}${b.row}`
}

interface SheetQualifier { text: string; linked: boolean }
const local: SheetQualifier = { text: '', linked: false }

function sheetPrefix(context: Context, index: number, tokens: Payload, xls: boolean): SheetQualifier | undefined {
  const xti = context.xti[index]
  if (!xti) return tokens.fail(`external sheet index ${index} out of bounds`)
  const [book, first, last] = xti
  const source = context.linkSources[book]
  if (!source) return tokens.fail(`external source index ${book} out of bounds`)
  if (first === -1 || last === -1) return { text: '#REF!', linked: false }
  const prefix = source.number === undefined ? '' : `[${source.number}]`
  if (xls && !source.self && source.number === undefined) return undefined
  const text = first === -2 ? `${prefix || '[0]'}!` : sheetSpan(source.sheets ?? context.sheets, first, last, prefix)
  return { text, linked: source.number !== undefined }
}

function sheetSpan(sheets: readonly string[], first: number, last: number, prefix = ''): string {
  const start = sheets[first]
  const end = sheets[last]
  if (start === undefined || end === undefined) return `${prefix}#REF!`
  if (first === last) return qualifier(prefix, start)
  const span = `${prefix}${start}:${end}`
  const plain = quoteSheetIfNeeded(start) === start && quoteSheetIfNeeded(end) === end
  return (plain ? span : quoteSheet(span)) + '!'
}

function binarySheetPrefix(context: Context, tokens: Payload, xls: XlsFormula | undefined): SheetQualifier | undefined {
  if (!xls?.old) return sheetPrefix(context, tokens.u16(), tokens, Boolean(xls))
  const link = tokens.i16()
  tokens.skip(8)
  const first = tokens.i16()
  const last = tokens.i16()
  if (link > 0) {
    const target = xls.links?.[link - 1]
    if (!target?.number || target.sheet === undefined) return undefined
    return { text: qualifier(`[${target.number}]`, target.sheet), linked: true }
  }
  if (xls.links?.[-link - 1]?.current) return { text: '!', linked: false }
  return { text: sheetSpan(context.sheets, first, last), linked: false }
}

function qualifier(prefix: string, sheet: string): string {
  const name = prefix + sheet
  return (quoteSheetIfNeeded(sheet) === sheet ? name : quoteSheet(name)) + '!'
}

function definedName(context: Context, index: number, tokens: Payload): Defined {
  return context.names[index - 1] ?? tokens.fail(`defined name index ${index} out of bounds`)
}

function internalName(context: Context, name: Defined): string {
  const sheet = context.sheets[name.scope]
  return (sheet === undefined ? '[0]!' : qualifier('', sheet)) + name.name
}

function externalName(number: number | undefined, name: string, sheet: string | undefined): string {
  if (number === undefined) return name
  const book = `[${number}]`
  return (sheet === undefined ? book + '!' : qualifier(book, sheet)) + name
}

function array(extra: Payload): string {
  const rows = extra.u32()
  const cols = extra.u32()
  if (!rows || !cols || rows > MAX_ROWS || cols > MAX_COLS || rows * cols * 2 > extra.remaining) extra.fail('invalid array constant dimensions')
  const elements: string[] = []
  for (let i = 0; i < rows * cols; i++) {
    const type = extra.u8()
    let text: string
    switch (type) {
      case 0: text = generalText(extra.f64(), 21); break
      case 1: text = textLiteral(extra.wide(extra.u16())); break
      case 2: text = extra.u8() ? 'TRUE' : 'FALSE'; break
      case 4: text = errorText(extra.u8()); extra.skip(3); break
      default: return extra.fail(`unknown array element type ${type}`)
    }
    elements.push(text)
  }
  const rendered: string[] = []
  for (let row = 0; row < rows; row++) rendered.push(elements.slice(row * cols, (row + 1) * cols).join(','))
  return '{' + rendered.join(';') + '}'
}

function list(tokens: Payload, extra: Payload, context: Context): string {
  const xti = tokens.u16()
  const flags = tokens.u16()
  const id = tokens.u32()
  const first = tokens.u16()
  const last = tokens.u16()
  if (flags & 0x1000) return '#REF!'
  if (flags & 0x2000) {
    const haveColumns = extra.u8()
    const areas = AREAS[extra.u16() & 31] ?? []
    const name = extra.wide(extra.u16())
    let cols: TableSpec['cols']
    if (haveColumns) {
      extra.u16()
      const count = extra.u8()
      if (count !== 1 && count !== 2) extra.fail(`invalid table column count ${count}`)
      extra.skip(3)
      const first = extra.wide(extra.u16())
      if (count === 2) {
        extra.skip(3)
        cols = [first, extra.wide(extra.u16())]
      } else cols = [first]
    }
    const source = context.xti[xti] && context.linkSources[context.xti[xti][0]]
    return source?.number ? `[${source.number}]!${linkedTableText({ name, spec: { areas, cols } }, true)}` : '#REF!'
  }
  const table = context.tables.get(id)
  if (!table) return tokens.fail(`table index ${id} out of bounds`)
  const n = flags & 3
  const areas = AREAS[(flags >>> 2) & 31]
  if (!areas || n > 2 || (n && first >= table.columns.length) || (n === 2 && last >= table.columns.length)) {
    return tokens.fail('invalid structured reference')
  }
  const cols: TableSpec['cols'] = n === 0 ? undefined : n === 1 ? [table.columns[first]] : [table.columns[first], table.columns[last]]
  return linkedTableText({ name: table.name, spec: { areas, cols } }, true)
}

export function formulaText({ tokens, extra, xls }: Formula, context: Context, row = 0, col = 0, relative = false, implicitSheet?: string): FormulaText {
  const biff2 = xls?.old === 2
  const narrowFunctions = biff2 || xls?.old === 3
  const stack: Operand[] = []
  let space = ''
  let openSpace = ''
  let closeSpace = ''
  const push = (text: string, name?: Defined) => {
    const value = space + text
    stack.push(name ? { text: value, name } : value)
    space = ''
  }
  const operands = (length: number) => {
    if (stack.length < length) tokens.fail(`formula token needs ${length} operands, has ${stack.length}`)
    return stack.splice(stack.length - length, length)
  }
  const args = (length: number) => operands(length).map(render)
  const parenthesized = (body: string) => {
    const text = '(' + body + closeSpace + ')'
    closeSpace = ''
    return text
  }
  while (tokens.remaining) {
    const byte = tokens.u8()
    const type = byte < 0x20 ? byte : 0x20 | (byte & 31)
    if (type >= 3 && type <= 0x11) {
      const [start, end] = operands(2)
      let left = render(start)
      let right = render(end)
      const gap = space
      space = ''
      if (type === 0x11) {
        if (typeof start !== 'string' && start.name.empty && right.trimStart().startsWith("'") && right.includes("'!")) {
          const name = start.name.name
          // An unbound R or C is a reserved name in legacy Excel. Conversion prefixes it with an underscore.
          const held = xls && /^[RC]$/i.test(name) ? '_' + name : quoteSheet(name)
          left = left.slice(0, -name.length) + held
        }
        // Excel quotes an explicit sheet after a range operator to distinguish it from a 3D sheet span.
        right = right.replace(/^((?:\[\d+\])?[\p{L}\p{N}_.]+)!/u, "'$1'!")
      }
      push(left + gap + BINARY[type - 3] + right)
      continue
    }
    switch (type) {
      case 1: case 2: {
        if (type === 2 && !xls) return undefined
        const anchorRow = xls ? tokens.u16() : tokens.u32()
        const anchorCol = biff2 ? tokens.u8() : xls ? tokens.u16() : extra.u32()
        if (xls && anchorCol > 255) return undefined
        const anchor = address(tokens, anchorRow, anchorCol)
        tokens.done()
        extra.done()
        return type === 1 ? { anchor } : { table: anchor }
      }
      case 0x12: case 0x13: push((type === 0x12 ? '+' : '-') + args(1)[0]); break
      case 0x14: push(args(1)[0] + '%'); break
      case 0x15:
        push(openSpace + parenthesized(args(1)[0]))
        openSpace = ''
        break
      case 0x16: push(''); break
      case 0x17: push(textLiteral(xls ? xls.string(tokens) : tokens.wide(tokens.u16()))); break
      case 0x18: {
        if (xls && (!tokens.remaining || (!xls.label && !context.pivotNames))) return undefined
        const extension = tokens.u8()
        if (extension === 0x19 && !xls) push(list(tokens, extra, context))
        else if (extension === 0x1d && context.pivotNames) {
          const index = tokens.u32()
          const name = context.pivotNames[index]
          if (name === undefined) return tokens.fail(`pivot formula name index ${index} out of bounds`)
          push(name)
        } else {
          const range = xls?.label?.(extension, tokens, extra, () => area(tokens, xls, false))
          if (range === undefined) return undefined
          push(range)
        }
        break
      }
      case 0x19: {
        const flag = tokens.u8()
        if (flag === 0x40 || flag === 0x41) {
          const position = tokens.u8()
          const count = tokens.u8()
          const padding = (position & 1 ? '\n' : ' ').repeat(count)
          if (position === 2 || position === 3) openSpace += padding
          else if (position === 4 || position === 5) closeSpace += padding
          else space += padding
        } else if (flag === 4) {
          const choices = biff2 ? tokens.u8() : tokens.u16()
          tokens.skip((biff2 ? 1 : 2) * (choices + 1))
        } else {
          tokens.skip(biff2 ? 1 : 2)
          if (flag === 0x10) push('SUM' + parenthesized(args(1)[0]))
        }
        break
      }
      case 0x1c: push(errorText(tokens.u8())); break
      case 0x1d: push(tokens.u8() ? 'TRUE' : 'FALSE'); break
      case 0x1e: push(String(tokens.u16())); break
      case 0x1f: {
        const value = tokens.f64()
        if (!Number.isFinite(value)) tokens.fail('non-finite number in formula')
        push(generalText(value, 21))
        break
      }
      case 0x20: {
        tokens.skip(biff2 ? 6 : xls ? 7 : 14)
        const constant = xls ? xls.array(extra) : array(extra)
        if (constant === undefined) return undefined
        push(constant)
        break
      }
      case 0x21: {
        const id = narrowFunctions ? tokens.u8() : tokens.u16()
        const fn = FTAB[id]
        if (!fn || fn[1] === undefined) return undefined
        push(fn[0] + parenthesized(args(fn[1]).join(',')))
        break
      }
      case 0x22: {
        const count = tokens.u8()
        const id = narrowFunctions ? tokens.u8() : tokens.u16()
        const values = operands(count)
        let fn = FTAB[id]?.[0]
        if (id === 255) {
          // With a union argument, Excel can put the future-function operand after earlier arguments.
          const index = values.findIndex((value) => typeof value !== 'string' && value.name.future)
          const [callee] = values.splice(Math.max(0, index), 1)
          fn = callee === undefined ? undefined : functionName(callee)
        }
        if (!fn) return undefined
        push(fn + parenthesized(values.map(render).join(',')))
        break
      }
      case 0x23: {
        const index = xls?.old ? tokens.u16() : tokens.u32()
        if (xls?.old) tokens.skip(biff2 ? 5 : xls.old < 5 ? 8 : 12)
        const name = definedName(context, index, tokens)
        push(name.name, name)
        break
      }
      case 0x24: case 0x2c: case 0x3a: {
        const owner = type === 0x3a ? binarySheetPrefix(context, tokens, xls) : local
        if (!owner) return undefined
        const prefix = owner.text
        const offset = type === 0x2c || (relative && type === 0x3a)
        const at = ref(tokens, xls ? tokens.u16() : tokens.u32(), xls?.old ? tokens.u8() : tokens.u16(), xls, offset ? row : undefined, offset ? col : undefined)
        push(xls && prefix.endsWith('#REF!') ? prefix : prefix + at.col + at.row)
        break
      }
      case 0x25: case 0x2d: case 0x3b: {
        const owner = type === 0x3b ? binarySheetPrefix(context, tokens, xls) : local
        if (!owner) return undefined
        const prefix = owner.text
        const offset = type === 0x2d || (relative && type === 0x3b)
        const range = area(tokens, xls, owner.linked, offset ? row : undefined, offset ? col : undefined)
        push(xls && prefix.endsWith('#REF!') ? prefix : prefix + range)
        break
      }
      case 0x26:
        tokens.skip(biff2 ? 4 : 6)
        extra.skip(xls ? extra.u16() * 8 : extra.u32() * 16)
        break
      case 0x27: case 0x28: tokens.skip(biff2 ? 4 : 6); break
      case 0x29: case 0x2e: case 0x2f: tokens.skip(biff2 ? 1 : 2); break
      case 0x2a: case 0x2b: case 0x3c: case 0x3d: {
        const owner = type >= 0x3c ? binarySheetPrefix(context, tokens, xls) : local
        if (!owner) return undefined
        const prefix = owner.text
        const width = xls?.old ? 3 : xls ? 4 : 6
        tokens.skip(type === 0x2a || type === 0x3c ? width : width * 2)
        // A conditional format in SpreadsheetML has its own sheet as the implicit reference.
        if (prefix.endsWith('#REF!')) push(prefix)
        else if (implicitSheet !== undefined && prefix === qualifier('', implicitSheet)) push('#REF!')
        else push(prefix + '#REF!')
        break
      }
      case 0x39: {
        if (xls?.old === 5) {
          const link = tokens.i16()
          tokens.skip(8)
          const index = tokens.u16()
          tokens.skip(12)
          if (link < 0) {
            push(internalName(context, definedName(context, index, tokens)))
          } else {
            const source = xls.links?.[link - 1]
            const name = source?.names?.[index - 1]
            if (!name) return undefined
            push(externalName(source?.number, name.name, name.sheet))
          }
          break
        }
        if (xls?.old) return undefined
        const link = tokens.u16()
        const index = tokens.u32()
        const xti = context.xti[link]
        if (!xti) tokens.fail(`external sheet index ${link} out of bounds`)
        const source = context.linkSources[xti[0]]
        if (xls && source?.self && !context.names[index - 1]) return undefined
        if (source?.self) {
          // Excel uses PtgNameX for explicitly qualified names in the same workbook.
          push(internalName(context, definedName(context, index, tokens)))
          break
        }
        const name = source?.names?.[index - 1]
        if (!name) return undefined
        const sheet = source.scopes?.[index - 1]
        const owner = sheet ? source.sheets?.[sheet - 1] : undefined
        push(externalName(source.number, name, owner))
        break
      }
      default: return undefined
    }
  }
  extra.done()
  if (stack.length !== 1) {
    if (!stack.length) return undefined
    tokens.fail(`formula left ${stack.length} operands`)
  }
  return { text: render(stack[0]) + space }
}
