/**
 * A sheet's data validation: `<dataValidations>`, and `<x14:dataValidations>` in the sheet's extensions, where Excel
 * writes a rule whose formulas refer to another sheet. Each rule is read as `DataValidation`, xlsx's defaults (false)
 * spelled out where they differ from the model's (true, as Excel makes a rule); a rule the model cannot hold (no range,
 * no formula its type needs) is left out, cells a rule shares with an earlier one go to the earlier one (which Excel shows:
 * the valid set `struct-held-2`), an IME mode is not kept, a list's items holding commas are split, and each is reported.
 */
import type { DataValidation, ValidationOperator, ValidationType } from '../../contract.js'
import { cornerOf, subtract, type Rect } from '../../structure/area.js'
import { FILE_TYPES, OPERATORS } from '../../validation/check.js'
import { COMPARED, rangeText, twoBounds } from '../../validation/rules.js'
import { formulaAt } from '../../validation/structure.js'
import { sheetKey } from '../../syntax.js'
import type { Issues, XlsxMessage } from '../issues.js'
import { attributes, element, elements, flag, ownEntry, type Element } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import { readFormula, type FormulaOptions } from './formulas.js'
import { sqrefRanges } from './refs.js'

/** The model's types by the names a file gives them. */
const TYPES = Object.fromEntries(Object.entries(FILE_TYPES).map(([model, file]) => [file, model])) as Record<string, ValidationType>

/** A rule as the file holds it: its attributes, its formulas' text and its range. */
interface RawRule {
  a: Record<string, string>
  formula1?: string
  formula2?: string
  sqref: string
}

const text = (e: Element | undefined) => (e?.body == null ? undefined : readText(e.body))

/**
 * The rules of `<dataValidations>`, before the sheet's extensions. A list whose item holds a comma (typed where the comma
 * is no separator) is the fallback of an `x12ac:list`, which splits it at the comma: reported.
 */
function plain(xml: string, sheet: string, issues: Issues): RawRule[] {
  return elements(element(xml, 'dataValidations')?.body ?? '', 'dataValidation').map((e) => {
    const a = attributes(e.attrs)
    const body = e.body ?? ''
    if (element(body, 'list')) issues.add('dataValidation', { key: 'ruleListCommas' }, { sheet, cell: a.sqref ?? '' })
    return { a, formula1: text(element(body, 'formula1')), formula2: text(element(body, 'formula2')), sqref: a.sqref ?? '' }
  })
}

/** The rules of Excel 2010's extension: formulas in `<xm:f>`, the range in `<xm:sqref>`. */
function extended(xml: string): RawRule[] {
  return elements(element(xml, 'dataValidations')?.body ?? '', 'dataValidation').map((e) => {
    const body = e.body ?? ''
    const f = (k: string) => text(element(element(body, k)?.body ?? '', 'f'))
    return { a: attributes(e.attrs), formula1: f('formula1'), formula2: f('formula2'), sqref: text(element(body, 'sqref')) ?? '' }
  })
}

/** The model's rule of a rule of the file, or why it cannot be one. */
function rule(raw: RawRule, sheet: string, issues: Issues, opts: FormulaOptions): DataValidation | XlsxMessage {
  const { a } = raw
  const type = ownEntry(TYPES, a.type ?? 'none')
  if (!type) return { key: 'ruleType', type: a.type }
  const out: DataValidation = { range: raw.sqref, type }
  const where = { sheet, cell: raw.sqref }
  // `between` is the default, left out
  if (COMPARED.has(type) && a.operator !== undefined && a.operator !== 'between' && OPERATORS.includes(a.operator as ValidationOperator)) out.operator = a.operator as ValidationOperator
  const formula = (t: string | undefined) => (t === undefined || t === '' ? undefined : readFormula(t, issues, where, opts).formula)
  if (type !== 'any') {
    const f1 = formula(raw.formula1)
    if (f1 === undefined) return { key: 'ruleNoFormula' }
    out.formula1 = f1
    if (twoBounds(out)) {
      const f2 = formula(raw.formula2)
      if (f2 === undefined) return { key: 'ruleNoSecond' }
      out.formula2 = f2
    }
  }
  // xlsx's defaults are false, the model's those of a rule Excel makes
  if (!flag(a.allowBlank)) out.allowBlank = false
  if (type === 'list' && flag(a.showDropDown)) out.dropdown = false
  for (const [k, file] of [['promptTitle', 'promptTitle'], ['prompt', 'prompt']] as const) if (a[file] !== undefined) out[k] = readText(a[file])
  if (!flag(a.showInputMessage)) out.showPrompt = false
  if (a.errorStyle === 'warning' || a.errorStyle === 'information') out.errorStyle = a.errorStyle
  for (const k of ['errorTitle', 'error'] as const) if (a[k] !== undefined) out[k] = readText(a[k])
  if (!flag(a.showErrorMessage)) out.showError = false
  if (a.imeMode !== undefined && a.imeMode !== 'noControl') issues.add('dataValidation', { key: 'ruleIme' }, where)
  return out
}

/**
 * The sheet's rules (`tail`: the worksheet after its cells, without its own extension list, whose content is `ext`),
 * none meeting an earlier one; a rule that loses its first cell so has its formulas written for its new first cell
 * (`sheets`: the book's sheets).
 */
export function readValidations(tail: string, ext: string, sheet: string, sheets: string[], issues: Issues, opts: FormulaOptions): DataValidation[] | undefined {
  const names = new Map(sheets.map((s) => [sheetKey(s), s]))
  const raws = [...plain(tail, sheet, issues), ...extended(ext)]
  const out: DataValidation[] = []
  const taken: Rect[] = []
  for (const raw of raws) {
    const where = { sheet, cell: raw.sqref }
    const r = rule(raw, sheet, issues, opts)
    if ('key' in r) {
      issues.add('dataValidation', r, where)
      continue
    }
    const parts = sqrefRanges(raw.sqref)
    const areas = parts.length && parts.every((x) => x.rect) ? parts.map((x) => x.rect!) : null
    if (!areas) {
      issues.add('dataValidation', { key: 'ruleNoRange' }, where)
      continue
    }
    const own = areas.flatMap((x) => taken.reduce(subtract, [x]))
    if (own.length !== areas.length || own.some((x, i) => x !== areas[i])) issues.add('dataValidation', { key: 'ruleShared' }, where)
    if (!own.length) continue
    taken.push(...own)
    const [from, to] = [cornerOf(areas), cornerOf(own)]
    for (const k of ['formula1', 'formula2'] as const) if (r[k] !== undefined) r[k] = formulaAt(r[k], sheet, from, to, names)
    out.push({ ...r, range: rangeText(own) })
  }
  return out.length ? out : undefined
}
