/**
 * A sheet's data validation as `<dataValidations>`, as Excel writes it: xlsx's defaults are false, so a rule writes the
 * model's defaults (true) as attributes; formulas in the file's spelling; texts cut to Excel's limits. A rule whose
 * formula the file cannot hold (it does not parse, or names a sheet, a table or a column the book does not have) is left
 * out. Both are reported.
 */
import type { DataValidation } from '../../contract.js'
import { FILE_TYPES } from '../../validation/check.js'
import { areasOf } from '../../validation/rules.js'
import { rectText } from '../../structure/area.js'
import { toFile } from '../formulas/to-file.js'
import type { Issues } from '../issues.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import { fileTree, type FileNames } from './formula.js'
import { clip } from './xml.js'

/** Excel's limits on a rule's texts: it does not load a file with longer ones. */
const LIMITS = { promptTitle: 32, prompt: 255, errorTitle: 32, error: 225 } as const

/** `sqref`: each range from its first cell to its last, whole columns and rows too, as Excel saves them (`O1:O1048576`, the set `valid/files`). */
const sqref = (range: string) => areasOf(range)!.map(rectText).join(' ')

export interface ValidationContext extends FileNames {
  sheet: string
  issues: Issues
}

/** The rules of a sheet as `<dataValidations>`; nothing without rules. */
export function validationsXml(rules: DataValidation[] | undefined, ctx: ValidationContext): string {
  const items = (rules ?? []).flatMap((v) => {
    const where = { sheet: ctx.sheet, cell: v.range }
    const formulas: string[] = []
    for (const k of ['formula1', 'formula2'] as const) {
      const text = v[k]
      if (text === undefined) continue
      const tree = fileTree(text, ctx)
      if (!tree) {
        ctx.issues.add('dataValidation', { key: 'ruleFormulaOut' }, where)
        return []
      }
      formulas.push(`<${k}>${escapeText(toFile(tree))}</${k}>`)
    }
    let attrs = v.type === 'any' ? '' : ` type="${FILE_TYPES[v.type]}"`
    if (v.errorStyle && v.errorStyle !== 'stop') attrs += ` errorStyle="${v.errorStyle}"`
    if (v.operator && v.operator !== 'between') attrs += ` operator="${v.operator}"`
    if (v.allowBlank !== false) attrs += ' allowBlank="1"'
    if (v.type === 'list' && v.dropdown === false) attrs += ' showDropDown="1"'
    if (v.showPrompt !== false) attrs += ' showInputMessage="1"'
    if (v.showError !== false) attrs += ' showErrorMessage="1"'
    for (const k of ['errorTitle', 'error', 'promptTitle', 'prompt'] as const) {
      const t = v[k]
      if (t !== undefined) attrs += ` ${k}="${escapeAttr(clip(t, LIMITS[k], ctx.issues, { key: 'ruleTextCut', field: k, max: LIMITS[k] }, where))}"`
    }
    return [`<dataValidation${attrs} sqref="${sqref(v.range)}">${formulas.join('')}</dataValidation>`]
  })
  return items.length ? `<dataValidations count="${items.length}">${items.join('')}</dataValidations>` : ''
}
