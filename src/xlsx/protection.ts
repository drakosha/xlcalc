/**
 * A sheet's protection (`SheetProtection`) as xlsx `<sheetProtection>` holds it. The file says the opposite of the
 * model: an attribute true protects what the model's field allows (`formatCells="0"` allows formatting cells), and an
 * attribute left out takes the schema's default, under which `objects` and `scenarios` are unprotected while Excel's
 * dialog protects them. A sheet is protected by `sheet="1"`; the password's attributes are not in the model.
 */
import type { SheetProtection } from '../contract.js'
import { SHEET_PROTECTION } from '../input/protection.js'
import { flag } from './xml/scan.js'

type Field = keyof SheetProtection

/** The fields in the order Excel writes their attributes, each with the schema's default: true where a missing attribute protects. */
const SCHEMA: readonly [Field, boolean][] = [
  ['objects', false],
  ['scenarios', false],
  ['formatCells', true],
  ['formatColumns', true],
  ['formatRows', true],
  ['insertColumns', true],
  ['insertRows', true],
  ['insertHyperlinks', true],
  ['deleteColumns', true],
  ['deleteRows', true],
  ['selectLockedCells', false],
  ['sort', true],
  ['autoFilter', true],
  ['pivotTables', true],
  ['selectUnlockedCells', false],
]

/** The attributes of a protection's password, which `writeKept` keeps as the file has them. */
const PASSWORD = ['password', 'algorithmName', 'hashValue', 'saltValue', 'spinCount']

/** The protection `<sheetProtection>`'s attributes give, none for a sheet they do not protect; `password`: they hold one. */
export function readProtection(a: Record<string, string>): { protection?: SheetProtection; password: boolean } {
  if (!flag(a.sheet)) return { password: false }
  const protection: SheetProtection = {}
  for (const [k, protects] of SCHEMA) {
    const allowed = !flag(a[k], protects)
    if (allowed !== SHEET_PROTECTION[k]) protection[k] = allowed
  }
  return { protection, password: PASSWORD.some((k) => a[k] !== undefined) }
}

/** The attributes of a protection, `sheet` first, each null where the schema's default says it: an element's own set by `setAttrs` keeps its password. */
export function protectionAttrs(p: SheetProtection): Record<string, string | null> {
  const out: Record<string, string | null> = { sheet: '1' }
  for (const [k, protects] of SCHEMA) {
    const protect = !(p[k] ?? SHEET_PROTECTION[k])
    out[k] = protect === protects ? null : protect ? '1' : '0'
  }
  return out
}

/** A sheet's `<sheetProtection>`, empty for a sheet without protection. */
export function protectionXml(p: SheetProtection | undefined): string {
  if (!p) return ''
  const attrs = Object.entries(protectionAttrs(p)).filter(([, v]) => v !== null)
  return `<sheetProtection${attrs.map(([k, v]) => ` ${k}="${v}"`).join('')}/>`
}
