/**
 * FILTERXML held out after `filterxml-held-2`, which showed a text node to be #VALUE! as the node itself (its value is
 * its children's text, and a text node has none, while an attribute holds its value as a child): the rule over text
 * nodes reached otherwise (the self axis, a union with elements, a predicate on their text), CDATA alone, attributes and
 * namespace nodes in a column with elements. Excel for Windows alone. Answers:
 * `test/excel/win/cases/filterxml-held-3.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

const F = 'FILTERXML'
const q = (s: string) => `"${s.replace(/"/g, '""')}"`
const fx = (xml: string, xpath: string) => `${F}(${q(xml)},${q(xpath)})`
const at = (xml: string, xpath: string, k: number) => `INDEX(${fx(xml, xpath)},${k})`

const DOC = '<a k="1"><b>2</b><![CDATA[3]]><c x="4">5</c></a>'

function build(): Probe[] {
  const { out, add } = collector()
  add('text', F, fx(DOC, '/a/b/text()'), fx(DOC, '//text()[.="2"]'), fx(DOC, '/a/b/text()/self::node()'), fx(DOC, '/a/text()'), fx(DOC, '//b/text()/..'), fx(DOC, '//*[text()="5"]'), at(DOC, '//b | //b/text()', 1), at(DOC, '//b | //b/text()', 2))
  add('attrs', F, at(DOC, '//@* | //c', 1), at(DOC, '//@* | //c', 2), at(DOC, '//@* | //c', 3), fx(DOC, '//c/@x/..'), fx(DOC, 'sum(//@*)'), fx(DOC, '//*[@x=4]'))
  add('ns', F, at('<a xmlns:p="u"><b>1</b></a>', '/a/namespace::*', 1), at('<a xmlns:p="u"><b>1</b></a>', '/a/namespace::*', 2), fx('<a xmlns:p="u"><b>1</b></a>', 'count(/a/namespace::*)'), fx('<a xmlns:p="u"><b>1</b></a>', '/a/namespace::p'))
  return out
}

const PROBES = build()

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./filterxml-held-3.predicted.json', import.meta.url), 'utf8'))
