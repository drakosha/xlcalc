/**
 * FILTERXML held out after `filterxml`, whose misses showed that a node's value is its own text (an element's text
 * children, trimmed, a CDATA section as written), a single node without text #VALUE! while several are counted each,
 * TRUE and FALSE as booleans, and names compared as written whatever their namespace. Around those rules: nodes without
 * text among others, text and CDATA together, which whitespace is trimmed, booleans and numbers as text, prefixed name
 * tests, the namespace axis, text nodes of mixed content, positions in unions and arrays of XML. Excel for Windows alone.
 * Answers: `test/excel/win/cases/filterxml-held.json`; `PREDICTED`, the engine's answers, was committed before Excel
 * ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

const F = 'FILTERXML'
const q = (s: string) => `"${s.replace(/"/g, '""')}"`
const fx = (xml: string, xpath: string) => `${F}(${q(xml)},${q(xpath)})`
const at = (xml: string, xpath: string, k: number) => `INDEX(${fx(xml, xpath)},${k})`
const rows = (xml: string, xpath: string) => `ROWS(${fx(xml, xpath)})`

const DOC = '<a id="1"><b k="v">one</b><b>two</b><!-- note --><c><d>4</d><d>5</d></c><?pi data?><e/></a>'
const NS = '<x:a xmlns:x="u" xmlns:y="w"><x:b>3</x:b><y:b>4</y:b><b>5</b></x:a>'

function build(): Probe[] {
  const { out, add } = collector()
  // nodes without text among others: each counts, its own value
  add('mixed', F, at(DOC, '//*', 1), at(DOC, '//*', 2), `ISERROR(${at(DOC, '//*', 4)})`, at(DOC, '//node()', 6), at(DOC, '//node()', 12), at(DOC, '//node()', 13), rows(DOC, '//c | //d'), at(DOC, '//c | //d', 2), `COUNTA(${fx(DOC, '//*')})`, `SUM(IFERROR(${fx(DOC, '//*')},0))`)
  // text and CDATA, whitespace of every kind
  for (const xml of ['<a> x <![CDATA[ y ]]> z </a>', '<a><![CDATA[a]]><![CDATA[b]]></a>', '<a><![CDATA[ ]]></a>', '<a>&#160;x</a>', '<a>&#32;x&#32;</a>', '<a>&#9;7&#10;</a>', '<a>x<b/>y</a>', '<a> <![CDATA[1]]> </a>', '<a>1<!-- c -->2</a>', '<a>1<?p?>2</a>'])
    add('text', F, fx(xml, '/a'))
  add('text', F, `LEN(${fx('<a>&#160;x</a>', '/a')})`, `CODE(${fx('<a>&#160;x</a>', '/a')})`, rows('<a>x<b/>y</a>', '/a/text()'), at('<a>x<b/>y</a>', '/a/text()', 2), fx("<a b='  x  '/>", '/a/@b'), fx('<a b="&#9;x&#10;"/>', '/a/@b'), `LEN(${fx('<a b="&#9;x&#10;"/>', '/a/@b')})`, fx('<a b="&lt;&amp;"/>', '/a/@b'))
  // booleans and numbers as text
  for (const t of [' true ', 'True', 'TRUE ', 'yes', 'FALSE', 'ИСТИНА', '1e400', '-0', ' -1.5e-3 ', '1 000', '15%', '2020/01/02', 'Jan 2 2020', '1:30 PM', '#N/A', '1E-400', '0.1E1'])
    add('values', F, fx(`<a>${t}</a>`, '/a'))
  add('values', F, `TYPE(${fx('<a>true</a>', '/a')})`, `ISTEXT(${fx('<a>1 000</a>', '/a')})`)
  // names as written: prefixes, wildcards over a prefix, name(), the namespace axis
  for (const x of ['//x:b', '//y:b', '//b', '//x:*', 'count(//x:*)', '//*[name()="y:b"]', '//*[local-name()="b"][3]', '//*[namespace-uri()="w"]', '/x:a/@*', '//z:b'])
    add('names', F, x.startsWith('//x:*') ? rows(NS, x) : fx(NS, x))
  add('names', F, rows(NS, '/x:a/namespace::*'), at(NS, '/x:a/namespace::*', 1), rows(NS, '//b/namespace::*'), fx('<a xmlns="u"><b k="1"/></a>', '//b/@k'), fx('<a xmlns:p="u"><p:b p:k="2"/></a>', '//@p:k'))
  // positions over unions, reverse axes, nested predicates
  const LIST = '<r><i>1</i><i>2</i><i>3</i><g><i>4</i><i>5</i></g></r>'
  for (const x of ['(//i | //g)[last()]', '//i[position()=last()-1]', '//i[3 > position()][last()]', '(//i)[last()]', '//g/i[1]/preceding::i[1]', '//g/i[2]/preceding::i[last()]', '//i[5 - position() = 3]', '(//g/i | //i[1])[2]', '//i[../g]', '//i[count(../i) = 2][1]', '//*[i > 4]/@*', '//i[following-sibling::i][last()]', '//i[not(following-sibling::*)][1]', '(//i)[position() mod 2 = 0][2]'])
    add('positions', F, fx(LIST, x))
  // arrays of XML, a column of them, errors in both arguments
  add('args', F, `${F}({"<a>1</a>";"<a>2</a>"},"/a")`, `INDEX(${F}({"<a>1</a>";"<a>2</a>"},"/a"),2)`, `INDEX(${F}({"<a>1</a>","<a>x</a>"},"/a"),2)`, `${F}(1/0,NA())`, `${F}("<a>1</a>",NA())`, `${F}("<a>1</a>","/a")*2`, `${F}("<a>1</a>","/a")&"!"`)
  return out
}

const PROBES = build()

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./filterxml-held.predicted.json', import.meta.url), 'utf8'))
