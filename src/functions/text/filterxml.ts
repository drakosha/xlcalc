/**
 * FILTERXML, this engine's own: the nodes an XPath expression selects in XML text (`../xml/`), each node's own text a
 * number, a boolean or text, one node a value and several a column (#VALUE! for a node without text). XML that is not
 * well formed, an expression that is not XPath or gives no nodes, and a selection of none are #VALUE!, as in Excel for
 * Windows.
 */
import { Arr } from '../../arrays/value.js'
import type { Cx, FnDef } from '../../args.js'
import { E, type V } from '../../values.js'
import { parseXml, type XNode } from '../xml/dom.js'
import { selectNodes } from '../xml/xpath.js'
import { capped } from './common.js'

const trim = (s: string) => s.replace(/^[ \t\n\r]+|[ \t\n\r]+$/g, '')

/**
 * A node's own text as Excel takes it: an element's text children joined, each trimmed but a CDATA section, any other
 * node's value trimmed; null for an element without text, a comment or a processing instruction.
 */
function ownText(n: XNode): string | null {
  if (n.kind === 'comment' || n.kind === 'pi' || n.kind === 'root') return null
  if (n.kind !== 'element') return trim(n.value)
  const s = n.children.map((c) => (c.kind !== 'text' ? '' : c.name ? c.value : trim(c.value))).join('')
  return s === '' ? null : s
}

/** A node's value: a number or a boolean its text reads as, else the text; #VALUE! for a node without text. */
function valueOf(cx: Cx, n: XNode): V {
  const s = ownText(n)
  if (s === null) return E.VALUE
  const v = cx.rules.valueText(s) ?? cx.rules.toBool(s)
  return typeof v === 'number' || typeof v === 'boolean' ? v : capped(s)
}

function filterxml(cx: Cx, xml: string, xpath: string): V | Arr {
  const doc = parseXml(xml)
  const nodes = doc && selectNodes(doc, xpath)
  if (!nodes?.length) return E.VALUE
  const v = nodes.map((n) => valueOf(cx, n))
  return v.length === 1 ? v[0] : new Arr(v.length, 1, v.length, 1, v)
}

export const functions: Record<string, FnDef> = {
  FILTERXML: { args: 's s', array: true, cx: true, fn: filterxml },
}
