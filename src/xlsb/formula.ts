import { quoteSheetIfNeeded } from '../a1.js'
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
  const name = operand.name.name.toUpperCase()
  // Excel writes unresolved calls that shadow worksheet functions as _xludf.
  if (!nativeFunctions.has(name) && !XLFN.has(name) && !XLWS.has(name)) return operand.text
  const space = operand.text.slice(0, operand.text.length - operand.name.name.length)
  return space + '_xludf.' + operand.name.name
}

export type FormulaText = { text: string } | { anchor: string } | undefined

export function formula(payload: Payload): Formula {
  const tokens = payload.section(payload.u32())
  const extra = payload.section(payload.u32())
  return { tokens, extra }
}

const BINARY = ['+', '-', '*', '/', '^', '&', '<', '<=', '=', '>=', '>', '<>', ' ', ',', ':']
const textLiteral = (value: string) => `"${value.replace(/"/g, '""')}"`
const AREAS: Readonly<Partial<Record<number, TableSpec['areas']>>> = {
  0: [], 1: ['all'], 2: ['headers'], 4: ['data'], 6: ['headers', 'data'], 8: ['totals'], 12: ['data', 'totals'], 16: ['row'],
}

/** Excel writes a formula's number literals as decimal in the range where its formula parser keeps 15 digits. */
function numeric(value: number): string {
  if (!value) return '0'
  const rounded = value.toPrecision(15).replace(/(?:\.0+|(?<=\.[0-9]*)0+)(?=e|$)/, '')
  const [digits, power] = rounded.split('e')
  if (power === undefined) return digits
  const exponent = Number(power)
  if (exponent > 20 || exponent < -308) return rounded.replace('e', 'E').replace(/E(\d)/, 'E+$1')
  const sign = digits[0] === '-' ? '-' : ''
  const unsigned = digits.slice(sign.length).replace('.', '')
  return sign + (exponent < 0 ? `0.${'0'.repeat(-exponent - 1)}${unsigned}` : unsigned.padEnd(exponent + 1, '0'))
}

/** The signed offset and the two relativity bits are separate for each coordinate. */
function coordinate(value: number, size: number, relative: boolean, origin?: number): number {
  const index = value & (size - 1)
  return relative && origin !== undefined ? (origin + index) & (size - 1) : index
}

function ref(payload: Payload, rowBits: number, colBits: number, row?: number, col?: number) {
  const relC = Boolean(colBits & 0x4000)
  const relR = Boolean(colBits & 0x8000)
  const r = coordinate(rowBits, MAX_ROWS, relR, row)
  const c = coordinate(colBits, MAX_COLS, relC, col)
  const a = address(payload, r, c)
  const rowText = String(r + 1)
  const letters = a.slice(0, -rowText.length)
  return { r, c, row: (relR ? '' : '$') + rowText, col: (relC ? '' : '$') + letters }
}

function area(payload: Payload, row?: number, col?: number): string {
  const r1 = payload.u32()
  const r2 = payload.u32()
  const c1 = payload.u16()
  const c2 = payload.u16()
  const a = ref(payload, r1, c1, row, col)
  const b = ref(payload, r2, c2, row, col)
  if (a.r === 0 && b.r === MAX_ROWS - 1) return `${a.col}:${b.col}`
  if (a.c === 0 && b.c === MAX_COLS - 1) return `${a.row}:${b.row}`
  return `${a.col}${a.row}:${b.col}${b.row}`
}

function sheetPrefix(context: Context, index: number, tokens: Payload): string {
  const xti = context.xti[index]
  if (!xti) return tokens.fail(`external sheet index ${index} out of bounds`)
  const [book, first, last] = xti
  const source = context.linkSources[book]
  if (!source) return tokens.fail(`external source index ${book} out of bounds`)
  if (first === -1 || last === -1) return '#REF!'
  const prefix = source.number === undefined ? '' : `[${source.number}]`
  if (first === -2) return `${prefix || '[0]'}!`
  const sheets = source.sheets ?? context.sheets
  const start = sheets[first]
  const end = sheets[last]
  if (start === undefined || end === undefined) return `${prefix}#REF!`
  if (first === last) return qualifier(prefix, start)
  return quoteSheetIfNeeded(start) === start && quoteSheetIfNeeded(end) === end
    ? `${prefix}${start}:${end}!`
    : `'${(prefix + start + ':' + end).replace(/'/g, "''")}'!`
}

function qualifier(prefix: string, sheet: string): string {
  return quoteSheetIfNeeded(sheet) === sheet
    ? `${prefix}${sheet}!`
    : `'${(prefix + sheet).replace(/'/g, "''")}'!`
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
      case 0: text = numeric(extra.f64()); break
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

export function formulaText({ tokens, extra }: Formula, context: Context, row = 0, col = 0, relative = false, implicitSheet?: string): FormulaText {
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
          // Excel quotes an unresolved name when the next sheet label already needs quotes.
          const name = start.name.name
          left = left.slice(0, -name.length) + quoteSheet(name)
        }
        // Excel quotes an explicit sheet after a range operator to distinguish it from a 3D sheet span.
        right = right.replace(/^((?:\[\d+\])?[\p{L}\p{N}_.]+)!/u, "'$1'!")
      }
      push(left + gap + BINARY[type - 3] + right)
      continue
    }
    switch (type) {
      case 1: {
        const anchor = address(tokens, tokens.u32(), extra.u32())
        tokens.done()
        extra.done()
        return { anchor }
      }
      case 0x12: case 0x13: push((type === 0x12 ? '+' : '-') + args(1)[0]); break
      case 0x14: push(args(1)[0] + '%'); break
      case 0x15:
        push(openSpace + '(' + args(1)[0] + closeSpace + ')')
        openSpace = ''
        closeSpace = ''
        break
      case 0x16: push(''); break
      case 0x17: push(textLiteral(tokens.wide(tokens.u16()))); break
      case 0x18: {
        const extension = tokens.u8()
        if (extension === 0x19) push(list(tokens, extra, context))
        else if (extension === 0x1d && context.pivotNames) {
          const index = tokens.u32()
          const name = context.pivotNames[index]
          if (name === undefined) return tokens.fail(`pivot formula name index ${index} out of bounds`)
          push(name)
        } else return undefined
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
          const choices = tokens.u16()
          tokens.skip(2 * (choices + 1))
        } else {
          tokens.u16()
          if (flag === 0x10) push(`SUM(${args(1)[0]})`)
        }
        break
      }
      case 0x1c: push(errorText(tokens.u8())); break
      case 0x1d: push(tokens.u8() ? 'TRUE' : 'FALSE'); break
      case 0x1e: push(String(tokens.u16())); break
      case 0x1f: {
        const value = tokens.f64()
        if (!Number.isFinite(value)) tokens.fail('non-finite number in formula')
        push(numeric(value))
        break
      }
      case 0x20:
        tokens.skip(14)
        push(array(extra))
        break
      case 0x21: {
        const id = tokens.u16()
        const fn = FTAB[id]
        if (!fn || fn[1] === undefined) return undefined
        push(`${fn[0]}(${args(fn[1]).join(',')})`)
        break
      }
      case 0x22: {
        const count = tokens.u8()
        const id = tokens.u16()
        const values = operands(count)
        let fn = FTAB[id]?.[0]
        if (id === 255) {
          // With a union argument, Excel can put the future-function operand after earlier arguments.
          const index = values.findIndex((value) => typeof value !== 'string' && value.name.future)
          const [callee] = values.splice(Math.max(0, index), 1)
          fn = callee === undefined ? undefined : functionName(callee)
        }
        if (!fn) return undefined
        push(`${fn}(${values.map(render).join(',')})`)
        break
      }
      case 0x23: {
        const index = tokens.u32()
        const name = context.names[index - 1]
        if (!name) tokens.fail(`defined name index ${index} out of bounds`)
        push(name.name, name)
        break
      }
      case 0x24: case 0x2c: case 0x3a: {
        const prefix = type === 0x3a ? sheetPrefix(context, tokens.u16(), tokens) : ''
        const offset = type === 0x2c || (relative && type === 0x3a)
        const at = ref(tokens, tokens.u32(), tokens.u16(), offset ? row : undefined, offset ? col : undefined)
        push(prefix + at.col + at.row)
        break
      }
      case 0x25: case 0x2d: case 0x3b: {
        const prefix = type === 0x3b ? sheetPrefix(context, tokens.u16(), tokens) : ''
        const offset = type === 0x2d || (relative && type === 0x3b)
        push(prefix + area(tokens, offset ? row : undefined, offset ? col : undefined))
        break
      }
      case 0x26:
        tokens.skip(6)
        extra.skip(extra.u32() * 16)
        break
      case 0x27: case 0x28: tokens.skip(6); break
      case 0x29: case 0x2e: case 0x2f: tokens.skip(2); break
      case 0x2a: case 0x2b: case 0x3c: case 0x3d: {
        const prefix = type >= 0x3c ? sheetPrefix(context, tokens.u16(), tokens) : ''
        tokens.skip(type === 0x2a || type === 0x3c ? 6 : 12)
        // A conditional format in SpreadsheetML has its own sheet as the implicit reference.
        if (prefix.endsWith('#REF!')) push(prefix)
        else if (implicitSheet !== undefined && prefix === qualifier('', implicitSheet)) push('#REF!')
        else push(prefix + '#REF!')
        break
      }
      case 0x39: {
        const link = tokens.u16()
        const index = tokens.u32()
        const xti = context.xti[link]
        if (!xti) tokens.fail(`external sheet index ${link} out of bounds`)
        const source = context.linkSources[xti[0]]
        const name = source?.names?.[index - 1]
        if (!name) return undefined
        const sheet = source.scopes?.[index - 1]
        const owner = sheet ? source.sheets?.[sheet - 1] : undefined
        const prefix = source.number === undefined ? '' : owner ? qualifier(`[${source.number}]`, owner) : `[${source.number}]!`
        push(prefix + name)
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
