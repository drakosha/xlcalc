/**
 * FILTERXML held out after `filterxml-held`, where the second text node of `<a>x<b/>y</a>` was #VALUE! while the
 * element's text was `xy`: text nodes after an element, a comment or a CDATA section, in one parent and in several, the
 * first and the later ones, against the element's own text. Excel for Windows alone. Answers:
 * `test/excel/win/cases/filterxml-held-2.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

const F = 'FILTERXML'
const q = (s: string) => `"${s.replace(/"/g, '""')}"`
const fx = (xml: string, xpath: string) => `${F}(${q(xml)},${q(xpath)})`
const at = (xml: string, xpath: string, k: number) => `INDEX(${fx(xml, xpath)},${k})`

function build(): Probe[] {
  const { out, add } = collector()
  for (const [xml, n] of [['<a>x<b/>y</a>', 2], ['<a>x<b/>y<c/>z</a>', 3], ['<a>x<!--c-->y</a>', 2], ['<a>x<![CDATA[y]]></a>', 2], ['<a>x<?p?>y</a>', 2], ['<a>x<b>1</b>y</a>', 2]] as const)
    for (let k = 1; k <= n; k++) add('siblings', F, at(xml, '/a/text()', k))
  add('single', F, fx('<a><b/>y</a>', '/a/text()'), fx('<a>x<b/></a>', '/a/text()'), fx('<a><b/>y</a>', '/a'), fx('<a><!--c-->y</a>', '/a/text()'), fx('<a>x<b/>y</a>', '/a/text()[2]'), fx('<a>x<b/>y</a>', '/a/b/following-sibling::text()'))
  for (let k = 1; k <= 3; k++) add('nodes', F, at('<r><i>1</i>t<i>2</i></r>', '/r/node()', k))
  for (let k = 1; k <= 4; k++) add('nodes', F, at('<a><b>one</b><b>two</b><c><d>4</d></c>t</a>', '//text()', k))
  add('nodes', F, at('<a><b>1</b><b>2</b></a>', '//b/text()', 2), at('<a><b>1</b><b>x<c/>2</b></a>', '//b/text()', 3), at('<a><b>1</b>s<b>2</b></a>', '/a/text() | //b', 2), at('<a><b>1</b>s<b>2</b></a>', '/a/text() | //b', 3))
  return out
}

const PROBES = build()

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./filterxml-held-2.predicted.json', import.meta.url), 'utf8'))
