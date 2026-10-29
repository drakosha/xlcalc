/**
 * FILTERXML, this engine's own: the nodes an XPath expression selects in XML text (`../xml/`), each node's text a value
 * as a cell reads typed text, one node a value and several a column. XML that is not well formed, an expression that is
 * not XPath or gives no nodes, and a selection of none are #VALUE!, as in Excel for Windows.
 */
import { Arr } from '../../arrays/value.js'
import type { Cx, FnDef } from '../../args.js'
import { E, type V } from '../../values.js'
import { parseXml, stringValue } from '../xml/dom.js'
import { selectNodes } from '../xml/xpath.js'
import { capped } from './common.js'

function filterxml(cx: Cx, xml: string, xpath: string): V | Arr {
  const doc = parseXml(xml)
  const nodes = doc && selectNodes(doc, xpath)
  if (!nodes?.length) return E.VALUE
  const v = nodes.map((n): V => {
    const s = stringValue(n)
    return cx.rules.valueText(s) ?? capped(s)
  })
  return v.length === 1 ? v[0] : new Arr(v.length, 1, v.length, 1, v)
}

export const functions: Record<string, FnDef> = {
  FILTERXML: { args: 's s', array: true, cx: true, fn: filterxml },
}
