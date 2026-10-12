/**
 * A sheet's conditional formats as Excel writes them: each rule a `<cfRule>` of a `<conditionalFormatting>` over its
 * ranges, its priority by its place in the list, its style a differential format of styles.xml (`dxfId`). Excel 2010's
 * rules go into the sheet's extension list (`x14:conditionalFormatting`): a data bar's 2010 part beside its 2007 rule,
 * joined by an id, and whole there the rules 2007 cannot hold (the icon sets 3Stars, 3Triangles and 5Boxes, custom
 * icons, formulas over other sheets), with their style inline (`x14:dxf`). A rule with a formula the file cannot hold
 * is left out and reported.
 */
import type { ConditionalFormat, ConditionalValue, SheetInput } from '../../contract.js'
import { ruleFormula } from '../../cond/rules.js'
import { sheetKey, walk } from '../../syntax/print.js'
import { toFile } from '../formulas/to-file.js'
import type { Issues } from '../issues.js'
import { X14_CF } from '../read/conditional.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import { fileTree, type FileNames } from './sheet.js'
import type { StyleWriter } from './styles.js'
import { rgb } from './xml.js'

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'
/** The extension of a 2007 data bar that names its 2010 part. */
const BAR_ID = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'
/** Icon sets 2007 does not have. */
const NEW_SETS = new Set(['3Stars', '3Triangles', '5Boxes'])
const TEXT_OPERATOR: Record<string, string> = { containsText: 'containsText', notContainsText: 'notContains', beginsWith: 'beginsWith', endsWith: 'endsWith' }

export interface ConditionalContext extends FileNames {
  styles: StyleWriter
  issues: Issues
}

export interface ConditionalWritten {
  /** The `<conditionalFormatting>` elements, after the merged areas. */
  main: string
  /** The extension of Excel 2010's rules (`<ext>`), for the sheet's extension list; empty without them. */
  ext: string
}

const hex = (n: number, w: number) => n.toString(16).toUpperCase().padStart(w, '0')

/** A formula as a file holds it, and whether it names another sheet; null when the file cannot hold it. */
function fileFormula(text: string, own: string, ctx: ConditionalContext): { text: string; other: boolean } | null {
  const ast = fileTree(text, ctx)
  if (!ast) return null
  let other = false
  walk(ast, (n) => {
    if ((n.k === 'ref' || n.k === 'name') && n.sheet && sheetKey(n.sheet) !== sheetKey(own)) other = true
  })
  return { text: toFile(ast), other }
}

const VALUE_2007: Record<string, string> = { autoMin: 'min', autoMax: 'max' }

/** Thrown by a formula the file cannot hold, which leaves its rule out. */
const UNFIT = Symbol('unfit')

export function conditionalXml(sheet: SheetInput, index: number, ctx: ConditionalContext): ConditionalWritten {
  const main: string[] = []
  const ext: string[] = []
  // the model's ranges as they are: Excel saves `A:A` as `A1:A1048576` and reads both (the set `cond/files`, case `wide`)
  const block = (ranges: string[], rules: string) => `<conditionalFormatting sqref="${escapeAttr(ranges.join(' '))}">${rules}</conditionalFormatting>`
  const block14 = (ranges: string[], rules: string) => `<x14:conditionalFormatting xmlns:xm="${XM}">${rules}<xm:sqref>${escapeText(ranges.join(' '))}</xm:sqref></x14:conditionalFormatting>`
  ;(sheet.conditionalFormats ?? []).map(withFormula).forEach((rule, i) => {
    const where = { sheet: sheet.name, cell: rule.ranges[0] }
    const id = `{${hex(index + 1, 8)}-0000-4000-8000-${hex(i + 1, 12)}}`
    // every formula of the rule as the file holds it
    let other = false
    const f = (text: string) => {
      const x = fileFormula(text, sheet.name, ctx)
      if (!x) throw UNFIT
      other ||= x.other
      return x.text
    }
    let own: string[]
    let values: ((v: ConditionalValue, x14: boolean) => string) | undefined
    try {
      own = rule.type === 'cellIs' ? rule.formulas.map(f) : 'formula' in rule && rule.formula !== undefined ? [f(rule.formula)] : []
      const vs = rule.type === 'colorScale' ? rule.points.map((p) => p.value) : rule.type === 'dataBar' ? [rule.min, rule.max] : rule.type === 'iconSet' ? rule.points.map((p) => p.value) : []
      const texts = vs.map((v) => (v.value === undefined ? undefined : f(v.value)))
      values = (v, x14) => {
        const t = texts[vs.indexOf(v)]
        const gte = rule.type === 'iconSet' && rule.points.find((p) => p.value === v)?.gte === false ? ' gte="0"' : ''
        if (x14) return `<x14:cfvo type="${v.type}"${gte}>${t === undefined ? '' : `<xm:f>${escapeText(t)}</xm:f>`}</x14:cfvo>`
        return `<cfvo type="${VALUE_2007[v.type] ?? v.type}"${t === undefined ? '' : ` val="${escapeAttr(t)}"`}${gte}/>`
      }
    } catch (e) {
      if (e !== UNFIT) throw e
      ctx.issues.add('conditionalFormat', { key: 'condFormulaOut' }, where)
      return
    }
    const priority = ` priority="${i + 1}"`
    const stop = rule.stopIfTrue ? ' stopIfTrue="1"' : ''
    const style = 'style' in rule && rule.style ? rule.style : undefined
    const x14 = other || (rule.type === 'iconSet' && (NEW_SETS.has(rule.icons) || rule.points.some((p) => p.icon !== undefined)))
    const formulaXml = (x: boolean) => own.map((t) => (x ? `<xm:f>${escapeText(t)}</xm:f>` : `<formula>${escapeText(t)}</formula>`)).join('')
    // the attributes of a rule's type and the content of its condition (2007 or 2010)
    const parts = (x: boolean): [string, string] => {
      switch (rule.type) {
        case 'cellIs':
          return [` operator="${rule.operator}"`, formulaXml(x)]
        case 'containsText':
        case 'notContainsText':
        case 'beginsWith':
        case 'endsWith':
          return [` operator="${TEXT_OPERATOR[rule.type]}" text="${escapeAttr(rule.text)}"`, formulaXml(x)]
        case 'timePeriod':
          return [` timePeriod="${rule.period}"`, formulaXml(x)]
        case 'top10':
          return [` rank="${rule.rank}"${rule.percent ? ' percent="1"' : ''}${rule.bottom ? ' bottom="1"' : ''}`, '']
        case 'aboveAverage':
          return [`${rule.below ? ' aboveAverage="0"' : ''}${rule.equal ? ' equalAverage="1"' : ''}${rule.stdDev ? ` stdDev="${rule.stdDev}"` : ''}`, '']
        case 'colorScale': {
          const tag = x ? 'x14:' : ''
          return ['', `<${tag}colorScale>${rule.points.map((p) => values!(p.value, x)).join('')}${rule.points.map((p) => `<${tag}color rgb="${rgb(p.color)}"/>`).join('')}</${tag}colorScale>`]
        }
        case 'iconSet': {
          const tag = x ? 'x14:' : ''
          const custom = x && rule.points.some((p) => p.icon !== undefined)
          const icons = custom ? rule.points.map((p, k) => (p.icon === null ? '<x14:cfIcon iconSet="NoIcons" iconId="0"/>' : `<x14:cfIcon iconSet="${(p.icon ?? { set: rule.icons }).set}" iconId="${p.icon?.index ?? (rule.reverse ? rule.points.length - 1 - k : k)}"/>`)).join('') : ''
          const attrs = `${rule.reverse ? ' reverse="1"' : ''}${rule.showValue === false ? ' showValue="0"' : ''}${custom ? ' custom="1"' : ''}`
          return ['', `<${tag}iconSet iconSet="${rule.icons}"${attrs}>${rule.points.map((p) => values!(p.value, x)).join('')}${icons}</${tag}iconSet>`]
        }
        case 'dataBar':
          return ['', '']
        default:
          return ['', formulaXml(x)]
      }
    }
    if (rule.type === 'dataBar') {
      const old = rule.excel2007 === true
      const lengths = old ? `${rule.minLength !== undefined && rule.minLength !== 10 ? ` minLength="${rule.minLength}"` : ''}${rule.maxLength !== undefined && rule.maxLength !== 90 ? ` maxLength="${rule.maxLength}"` : ''}` : ''
      const link = old ? '' : `<extLst><ext uri="${BAR_ID}" xmlns:x14="${X14}"><x14:id>${id}</x14:id></ext></extLst>`
      const bar = `<dataBar${lengths}${rule.showValue === false ? ' showValue="0"' : ''}>${values!(rule.min, false)}${values!(rule.max, false)}<color rgb="${rgb(rule.color)}"/></dataBar>`
      main.push(block(rule.ranges, `<cfRule type="dataBar"${priority}${stop}>${bar}${link}</cfRule>`))
      if (old) return
      const a = [
        ` minLength="${rule.minLength ?? 0}" maxLength="${rule.maxLength ?? 100}"`,
        rule.gradient === false ? ' gradient="0"' : '',
        rule.border ? ' border="1"' : '',
        rule.direction ? ` direction="${rule.direction}"` : '',
        rule.negativeColor === undefined ? ' negativeBarColorSameAsPositive="1"' : '',
        rule.border && rule.negativeBorder !== undefined ? ' negativeBarBorderColorSameAsPositive="0"' : '',
        rule.axis ? ` axisPosition="${rule.axis}"` : '',
      ].join('')
      const colors = [
        rule.border ? `<x14:borderColor rgb="${rgb(rule.border)}"/>` : '',
        rule.negativeColor ? `<x14:negativeFillColor rgb="${rgb(rule.negativeColor)}"/>` : '',
        rule.border && rule.negativeBorder ? `<x14:negativeBorderColor rgb="${rgb(rule.negativeBorder)}"/>` : '',
        `<x14:axisColor rgb="${rgb(rule.axisColor ?? '#000000')}"/>`,
      ].join('')
      ext.push(block14(rule.ranges, `<x14:cfRule type="dataBar" id="${id}"><x14:dataBar${a}>${values!(rule.min, true)}${values!(rule.max, true)}${colors}</x14:dataBar></x14:cfRule>`))
      return
    }
    const [attrs, content] = parts(x14)
    if (x14) {
      const dxf = style ? `<x14:dxf>${ctx.styles.dxfBody(ctx.styles.dxf(style))}</x14:dxf>` : ''
      ext.push(block14(rule.ranges, `<x14:cfRule type="${rule.type}"${priority}${stop}${attrs} id="${id}">${content}${dxf}</x14:cfRule>`))
      return
    }
    const dxf = style ? ` dxfId="${ctx.styles.dxf(style)}"` : ''
    main.push(block(rule.ranges, `<cfRule type="${rule.type}"${dxf}${priority}${stop}${attrs}>${content}</cfRule>`))
  })
  return { main: main.join(''), ext: ext.length ? `<ext uri="${X14_CF}" xmlns:x14="${X14}"><x14:conditionalFormattings>${ext.join('')}</x14:conditionalFormattings></ext>` : '' }
}

/** A rule with the formula it computes by: a rule Excel computes by one of its own (`defaultFormula`: text, blanks, errors, a period) always has one in a file. */
export function withFormula(rule: ConditionalFormat): ConditionalFormat {
  if ('formula' in rule) return rule
  const formula = ruleFormula(rule)
  return formula === undefined ? rule : ({ ...rule, formula } as ConditionalFormat)
}
