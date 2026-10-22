/**
 * Conditional formats checked as input (`SheetInput.conditionalFormats`, a layout edit of them): the fields each type
 * of rule takes, its ranges on the sheet, thresholds and colours. A formula is not parsed here: one that does not
 * parse never holds, as a cell's formula computes an error.
 */
import type { ConditionalFormat, IconSetName } from '../contract.js'
import { InputError } from '../input-error.js'
import { array, checkBoolean, checkNumber, checkOneOf, checkString, object, onlyKeys } from '../input/basic.js'
import { checkRange } from '../input/range.js'
import { checkBorder, checkColor, checkFill, checkFont, checkFormat, type Given } from '../input/styles.js'

/** The icon sets by name, with their number of icons. */
export const ICON_SETS: Record<IconSetName, number> = {
  '3Arrows': 3,
  '3ArrowsGray': 3,
  '3Flags': 3,
  '3TrafficLights1': 3,
  '3TrafficLights2': 3,
  '3Signs': 3,
  '3Symbols': 3,
  '3Symbols2': 3,
  '3Stars': 3,
  '3Triangles': 3,
  '4Arrows': 4,
  '4ArrowsGray': 4,
  '4RedToBlack': 4,
  '4Rating': 4,
  '4TrafficLights': 4,
  '5Arrows': 5,
  '5ArrowsGray': 5,
  '5Rating': 5,
  '5Quarters': 5,
  '5Boxes': 5,
}

export const CELL_IS = ['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween'] as const
export const PERIODS = ['today', 'yesterday', 'tomorrow', 'last7Days', 'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth'] as const
export const VALUE_TYPES = ['min', 'max', 'num', 'percent', 'percentile', 'formula', 'autoMin', 'autoMax']
export const WITH_VALUE: ReadonlySet<string> = new Set(['num', 'percent', 'percentile', 'formula'])

const BASE = ['type', 'ranges', 'stopIfTrue']

/** The fields of each type of rule, besides `type`, `ranges` and `stopIfTrue`. */
const FIELDS: Record<ConditionalFormat['type'], string[]> = {
  cellIs: ['operator', 'formulas', 'style'],
  expression: ['formula', 'style'],
  containsText: ['text', 'formula', 'style'],
  notContainsText: ['text', 'formula', 'style'],
  beginsWith: ['text', 'formula', 'style'],
  endsWith: ['text', 'formula', 'style'],
  containsBlanks: ['formula', 'style'],
  notContainsBlanks: ['formula', 'style'],
  containsErrors: ['formula', 'style'],
  notContainsErrors: ['formula', 'style'],
  timePeriod: ['period', 'formula', 'style'],
  top10: ['rank', 'percent', 'bottom', 'style'],
  aboveAverage: ['below', 'equal', 'stdDev', 'style'],
  duplicateValues: ['style'],
  uniqueValues: ['style'],
  colorScale: ['points'],
  dataBar: ['min', 'max', 'color', 'showValue', 'minLength', 'maxLength', 'gradient', 'border', 'direction', 'negativeColor', 'negativeBorder', 'axis', 'axisColor', 'excel2007'],
  iconSet: ['icons', 'points', 'reverse', 'showValue'],
}

/** The types of rules with a style (a differential format): all but colour scales, data bars and icon sets. */
export const STYLED: ReadonlySet<string> = new Set(Object.entries(FIELDS).flatMap(([t, fields]) => (fields.includes('style') ? [t] : [])))

/** A field of a conditional style is given unless undefined, as a cell style's is (`checkStyle` of `input.ts`). */
const given: Given = (o, k) => o[k] !== undefined

/** A rule's style: a cell's style's font, fill and border, without a font's name and size (a dxf has none) and with a number format. */
function checkStyle(v: unknown, p: string) {
  const s = object(v, p, 'a style such as { fill: { color: "#FFC7CE" }, font: { color: "#9C0006" } }')
  onlyKeys(s, p, ['font', 'fill', 'border', 'numberFormat'])
  if (given(s, 'font')) checkFont(s.font, `${p}.font`, given, false)
  if (given(s, 'fill')) checkFill(s.fill, `${p}.fill`, '#FFC7CE')
  if (given(s, 'border')) checkBorder(s.border, `${p}.border`, given)
  if (s.numberFormat !== undefined) checkFormat(s.numberFormat, `${p}.numberFormat`)
}

function checkValue(v: unknown, p: string) {
  const o = object(v, p, 'a threshold such as { type: "min" } or { type: "percent", value: "50" }')
  onlyKeys(o, p, ['type', 'value'])
  checkOneOf(o.type, `${p}.type`, VALUE_TYPES)
  if (WITH_VALUE.has(o.type as string)) checkString(o.value, `${p}.value`)
  else if (o.value !== undefined) throw new InputError(`${p}.value`, `a threshold of type ${o.type} takes no value`)
}

function checkIcon(v: unknown, p: string) {
  if (v === null) return
  const o = object(v, p, 'an icon such as { set: "3Flags", index: 0 }, or null for none')
  onlyKeys(o, p, ['set', 'index'])
  checkOneOf(o.set, `${p}.set`, Object.keys(ICON_SETS))
  checkNumber(o.index, `${p}.index`, 0, ICON_SETS[o.set as IconSetName] - 1, true)
}

/** One rule of a sheet's conditional formats. */
function checkRule(v: unknown, p: string): ConditionalFormat {
  const o = object(v, p, 'a conditional format such as { type: "cellIs", ranges: ["B2:B9"], operator: "greaterThan", formulas: ["100"], style: { fill: { color: "#FFC7CE" } } }')
  checkOneOf(o.type, `${p}.type`, Object.keys(FIELDS))
  const type = o.type as ConditionalFormat['type']
  onlyKeys(o, p, [...BASE, ...FIELDS[type]])
  const ranges = array(o.ranges, `${p}.ranges`, 'an array of ranges such as ["B2:B9"]')
  if (!ranges.length) throw new InputError(`${p}.ranges`, 'a conditional format applies to at least one range')
  ranges.forEach((r, i) => checkRange(r, `${p}.ranges[${i}]`, false))
  if (o.stopIfTrue !== undefined) checkBoolean(o.stopIfTrue, `${p}.stopIfTrue`)
  if (o.style !== undefined) checkStyle(o.style, `${p}.style`)
  if (o.formula !== undefined) checkString(o.formula, `${p}.formula`)
  switch (type) {
    case 'cellIs': {
      checkOneOf(o.operator, `${p}.operator`, CELL_IS)
      const n = o.operator === 'between' || o.operator === 'notBetween' ? 2 : 1
      const fs = array(o.formulas, `${p}.formulas`, 'an array of formulas such as ["100"]')
      if (fs.length !== n) throw new InputError(`${p}.formulas`, `${o.operator} takes ${n} formula${n > 1 ? 's' : ''}, got ${fs.length}`)
      fs.forEach((f, i) => checkString(f, `${p}.formulas[${i}]`))
      break
    }
    case 'expression':
      checkString(o.formula, `${p}.formula`)
      break
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      checkString(o.text, `${p}.text`)
      break
    case 'timePeriod':
      checkOneOf(o.period, `${p}.period`, PERIODS)
      break
    case 'top10':
      checkNumber(o.rank, `${p}.rank`, 1, o.percent === true ? 100 : 1000, true)
      for (const k of ['percent', 'bottom']) if (o[k] !== undefined) checkBoolean(o[k], `${p}.${k}`)
      break
    case 'aboveAverage':
      for (const k of ['below', 'equal']) if (o[k] !== undefined) checkBoolean(o[k], `${p}.${k}`)
      if (o.stdDev !== undefined) checkNumber(o.stdDev, `${p}.stdDev`, 1, 3, true)
      if (o.stdDev !== undefined && o.equal === true) throw new InputError(`${p}.equal`, 'a rule over standard deviations takes no equal (Excel keeps one or the other)')
      break
    case 'colorScale': {
      const pts = array(o.points, `${p}.points`, 'an array of two or three points such as [{ value: { type: "min" }, color: "#F8696B" }, { value: { type: "max" }, color: "#63BE7B" }]')
      if (pts.length < 2 || pts.length > 3) throw new InputError(`${p}.points`, `a color scale has two or three points, got ${pts.length}`)
      pts.forEach((x, i) => {
        const pp = `${p}.points[${i}]`
        const pt = object(x, pp, 'a point such as { value: { type: "min" }, color: "#F8696B" }')
        onlyKeys(pt, pp, ['value', 'color'])
        checkValue(pt.value, `${pp}.value`)
        checkColor(pt.color, `${pp}.color`)
      })
      break
    }
    case 'dataBar':
      checkValue(o.min, `${p}.min`)
      checkValue(o.max, `${p}.max`)
      checkColor(o.color, `${p}.color`)
      for (const k of ['showValue', 'gradient', 'excel2007']) if (o[k] !== undefined) checkBoolean(o[k], `${p}.${k}`)
      for (const k of ['minLength', 'maxLength']) if (o[k] !== undefined) checkNumber(o[k], `${p}.${k}`, 0, 100, true)
      if ((o.minLength as number | undefined) !== undefined && (o.maxLength as number | undefined) !== undefined && (o.minLength as number) > (o.maxLength as number))
        throw new InputError(`${p}.minLength`, `the shortest bar (${o.minLength}) is longer than the longest (${o.maxLength})`)
      for (const k of ['border', 'negativeColor', 'negativeBorder', 'axisColor']) if (o[k] !== undefined) checkColor(o[k], `${p}.${k}`)
      if (o.direction !== undefined) checkOneOf(o.direction, `${p}.direction`, ['leftToRight', 'rightToLeft'])
      if (o.axis !== undefined) checkOneOf(o.axis, `${p}.axis`, ['middle', 'none'])
      break
    case 'iconSet': {
      checkOneOf(o.icons, `${p}.icons`, Object.keys(ICON_SETS))
      const n = ICON_SETS[o.icons as IconSetName]
      const pts = array(o.points, `${p}.points`, `an array of ${n} thresholds`)
      if (pts.length !== n) throw new InputError(`${p}.points`, `the set ${o.icons} has ${n} icons, so ${n} thresholds; got ${pts.length}`)
      pts.forEach((x, i) => {
        const pp = `${p}.points[${i}]`
        const pt = object(x, pp, 'a threshold such as { value: { type: "percent", value: "33" } }')
        onlyKeys(pt, pp, ['value', 'gte', 'icon'])
        checkValue(pt.value, `${pp}.value`)
        if (pt.gte !== undefined) checkBoolean(pt.gte, `${pp}.gte`)
        if (pt.icon !== undefined) checkIcon(pt.icon, `${pp}.icon`)
      })
      for (const k of ['reverse', 'showValue']) if (o[k] !== undefined) checkBoolean(o[k], `${p}.${k}`)
      break
    }
  }
  return o as unknown as ConditionalFormat
}

/** A sheet's conditional formats. */
export function checkConditionalFormats(v: unknown, path: string): ConditionalFormat[] {
  return array(v, path, 'an array of conditional formats').map((r, i) => checkRule(r, `${path}[${i}]`))
}
