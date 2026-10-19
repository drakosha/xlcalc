/**
 * Checks of a sheet's data validation and outline settings at the API boundary (`InputError`): each rule's fields, a
 * range of the sheet's cells, the formulas its type takes, and no two rules over one cell.
 */
import { overlapping } from '../layout/merges.js'
import { InputError } from '../input-error.js'
import { checkBoolean, checkOneOf, checkString, object, onlyKeys } from '../input/basic.js'
import type { Rect } from '../structure/area.js'
import { areasOf, COMPARED, twoBounds } from './rules.js'
import type { DataValidation, ValidationOperator, ValidationType } from '../contract.js'

/** The types of rules with the names an xlsx file gives them: `any` is the file's `none`. */
export const FILE_TYPES: Readonly<Record<ValidationType, string>> = { any: 'none', whole: 'whole', decimal: 'decimal', date: 'date', time: 'time', textLength: 'textLength', list: 'list', custom: 'custom' }
const TYPES = Object.keys(FILE_TYPES)
/** The operators of the compared types (`COMPARED`), `between` by default. */
export const OPERATORS: readonly ValidationOperator[] = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']
const FIELDS = ['range', 'type', 'operator', 'formula1', 'formula2', 'allowBlank', 'dropdown', 'promptTitle', 'prompt', 'showPrompt', 'errorStyle', 'errorTitle', 'error', 'showError']

/** The rules of a sheet: each well formed, none meeting another. */
export function checkValidations(v: unknown, path: string): DataValidation[] {
  if (!Array.isArray(v)) throw new InputError(path, 'expected an array of rules such as [{ range: "B2:B9", type: "list", formula1: \'"Yes,No"\' }]')
  const areas: Rect[] = []
  const owner: [rule: number, text: string][] = []
  v.forEach((x, i) => {
    const p = `${path}[${i}]`
    const o = object(x, p, 'a rule such as { range: "B2:B9", type: "whole", operator: "greaterThan", formula1: "0" }')
    onlyKeys(o, p, FIELDS)
    const xs = areasOf(checkString(o.range, `${p}.range`))
    if (!xs) throw new InputError(`${p}.range`, `expected ranges of the sheet separated by spaces, such as "B2:B9 D2", got ${JSON.stringify(o.range)}`)
    for (const a of xs) {
      areas.push(a)
      owner.push([i, o.range as string])
    }
    checkOneOf(o.type, `${p}.type`, TYPES)
    const rule = o as unknown as DataValidation
    if (o.operator !== undefined) {
      if (!COMPARED.has(rule.type)) throw new InputError(`${p}.operator`, `a rule of type ${rule.type} has no operator`)
      checkOneOf(o.operator, `${p}.operator`, OPERATORS)
    }
    for (const k of ['formula1', 'formula2', 'promptTitle', 'prompt', 'errorTitle', 'error']) if (o[k] !== undefined) checkString(o[k], `${p}.${k}`)
    if (rule.type !== 'any' && o.formula1 === undefined) throw new InputError(`${p}.formula1`, `a rule of type ${rule.type} needs formula1`)
    if (rule.type === 'any' && o.formula1 !== undefined) throw new InputError(`${p}.formula1`, 'a rule of type any has no formula')
    if (twoBounds(rule) !== (o.formula2 !== undefined)) throw new InputError(`${p}.formula2`, twoBounds(rule) ? 'between and notBetween need formula2' : 'only between and notBetween take formula2')
    for (const k of ['allowBlank', 'dropdown', 'showPrompt', 'showError']) if (o[k] !== undefined) checkBoolean(o[k], `${p}.${k}`)
    if (o.errorStyle !== undefined) checkOneOf(o.errorStyle, `${p}.errorStyle`, ['stop', 'warning', 'information'])
  })
  const o = overlapping(areas)
  if (o) throw new InputError(`${path}[${owner[o[1]][0]}].range`, `${owner[o[1]][1]} meets ${owner[o[0]][1]}: a cell has one rule at most`)
  return v as DataValidation[]
}

/** Outline settings: where summary rows and columns are. */
export function checkOutline(v: unknown, path: string) {
  const o = object(v, path, 'outline settings such as { summaryBelow: false }')
  onlyKeys(o, path, ['summaryBelow', 'summaryRight'])
  for (const k of ['summaryBelow', 'summaryRight']) if (o[k] !== undefined) checkBoolean(o[k], `${path}.${k}`)
}
