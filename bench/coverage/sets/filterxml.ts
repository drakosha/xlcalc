/**
 * FILTERXML, which Excel for Windows alone has (Excel for Mac 16.114 does not know it): the values of the nodes it
 * selects (numbers, text that reads as other values, whitespace, mixed content), XPath 1.0 (axes, predicates, the core
 * functions, unions, expressions that give no nodes), XML that is not well formed, namespaces, and arrays given for
 * either argument. Answers: `test/excel/win/cases/filterxml.json`; `PREDICTED`, the engine's answers, was committed
 * before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

const F = 'FILTERXML'
const q = (s: string) => `"${s.replace(/"/g, '""')}"`
const fx = (xml: string, xpath: string) => `${F}(${q(xml)},${q(xpath)})`
/** The k-th of the nodes, so that several do not spill. */
const at = (xml: string, xpath: string, k: number) => `INDEX(${fx(xml, xpath)},${k})`
const rows = (xml: string, xpath: string) => `ROWS(${fx(xml, xpath)})`

const LIST = '<r><i>1</i><i>2</i><i>x</i><i k="a">3</i><g><i>4</i></g></r>'
const DOC = '<a id="1"><b k="v" n="2">one</b><b>two</b><!-- note --><c><d>4</d><d>5</d></c><?pi data?><e/></a>'

/** Many nodes in one text: 300 elements. */
const MANY = `<r>${Array.from({ length: 300 }, (_, i) => `<i>${i}</i>`).join('')}</r>`

/** Sheet X: XML in cells, a text longer than a formula's literal may hold. */
export const SHEETS: SheetCells = { X: { A1: '<a><b>1</b><b>2</b></a>', A2: '//b', A3: 5, A4: '<a>7</a>', A5: MANY, A6: `<a>${'z'.repeat(300)}</a>` } }

function build(): Probe[] {
  const { out, add } = collector()
  // the nodes of one document: elements, attributes, text, comments, processing instructions; one and several
  add('nodes', F, fx(DOC, '//b[1]'), fx(DOC, '/a/b[2]'), at(DOC, '//b', 2), rows(DOC, '//b'), fx(DOC, '//b/@k'), fx(DOC, '//@n'), fx(DOC, '/a/@id'), rows(DOC, '//@*'))
  add('nodes', F, fx(DOC, '//comment()'), fx(DOC, '//processing-instruction()'), fx(DOC, '//c'), fx(DOC, '/a'), fx(DOC, '/'), fx(DOC, '//e'), fx(DOC, '//b/text()'), rows(DOC, '//node()'), rows(DOC, '//*'), rows(DOC, '//text()'))
  // values: numbers, text that reads as other values, empty and white text
  for (const t of ['1', ' 2 ', '1.5', '-3', '1e3', '1E+3', '1,000', '$5', '50%', 'TRUE', 'false', '1/2/2020', '12:30', '2020-01-02', '0x10', '(4)', '.5', '1.', '+7', '00012', '1 2', ''])
    add('values', F, fx(`<a>${t}</a>`, '/a'))
  add('values', F, fx('<a> </a>', '/a'), fx('<a>  x  y  </a>', '/a'), fx('<a>x\ny</a>', '/a'), `LEN(${fx('<a> x </a>', '/a')})`, `ISTEXT(${fx('<a>1</a>', '/a')})`, `ISNUMBER(${fx('<a> 1 </a>', '/a')})`)
  add('values', F, fx('<a>x<b>1</b>y</a>', '/a'), fx('<a> <b>1</b> </a>', '/a/text()'), fx('<a><b>1</b> <b>2</b></a>', '/a'), fx('<a>x <b>1</b></a>', '/a'), fx('<a><![CDATA[ c ]]></a>', '/a'), fx('<a b=" 1 "/>', '/a/@b'), fx('<a b="x&#10;y"/>', '/a/@b'))
  add('values', F, fx('<a>&lt;&amp;&gt;&quot;&apos;</a>', '/a'), fx('<a>&#65;&#x42;</a>', '/a'), fx('<a>&#160;1</a>', '/a'), `LEN(${F}(X!A6,"/a"))`)
  // XPath: predicates, axes, operators, functions
  for (const x of ['//i[2]', '(//i)[5]', '//i[last()]', '//i[position()<3][2]', '//i[.>1]', '//i[.="x"]', '//i[@k]', '//i[not(@k)][3]', '//i[@k="a"]', '//g/i', '//g/..', '//i[4]/following-sibling::*', '//i[3]/preceding-sibling::i[1]', '//g/i/ancestor::*[last()]', '//i/parent::g', '//i[contains(.,"x")]', '//i[starts-with(.,"3")]', '//i[string-length(.)=1][4]', '//i[number(.)=4]', '//*[count(i)>1]', '//*[sum(i)=6]', '/r/*[name()="g"]', '//i[. mod 2 = 1][2]', '//i[.*2 = 4]', '//i[. div 2 = 1.5]', '//i[-. = -2]', '//i[.=2 or .=4][2]', '//i[.>1 and .<4]', '//i[normalize-space(.)="x"]', '//i[translate(.,"x","y")="y"]', '//i[substring(.,1,1)="3"]', '//i[concat(.,"!")="2!"]', '//i[floor(.)=2]', '//i[round(.)=4]', '//i[boolean(@k)]', '//i[true()][1]', '//i[1] | //g', '//g | //i[1]', 'r/i[1]', './/i[2]', '//self::g', '//i[1]/self::i', '/descendant::i[3]', '/descendant-or-self::g', '//i[1]/following::i[1]', '//g/preceding::i[1]', '//i[@k]/@k', '//@*/..', '//i[lang("en")]', '//*[local-name()="g"]'])
    add('xpath', F, x.includes('|') || x.includes('/*') || x.includes('/@*') ? rows(LIST, x) : fx(LIST, x))
  // expressions that give no node-set, errors of the grammar, nothing selected
  for (const x of ['count(//i)', 'string(//i)', '1+1', 'true()', '"x"', '//i[1]/text()=1', '', '//', '///i', '//i[', '//i]', '//i[1', '$v', 'p:i', '//p:*', 'foo()', 'count()', '//i[@]', '//zz', '//i[9]', '/r/i/i', 'child::', 'bogus::i', '//i/namespace::*', '//I'])
    add('xpath-off', F, fx(LIST, x))
  // XML that is not well formed, a declaration, a DTD, comments and instructions around the root
  for (const xml of ['<a>1', '<a>1</b>', '<a/><b/>', 'x<a/>', '<a/>x', '<a>&foo;</a>', '<a>&</a>', '<a b=1/>', '<a b="1" b="2"/>', '<!DOCTYPE a><a>1</a>', '<?xml version="1.0"?><a>1</a>', '<?xml version="1.0" encoding="UTF-8"?><a>1</a>', ' <a>1</a> ', '<!-- c --><a>1</a><?p x?>', '<a>]]></a>', '<a><!-- x -- y --></a>', '<a>&#0;</a>', '<a>&#x110000;</a>', '<1a/>', '<a:b:c/>', '', 'text', '<a><b></a></b>', '<a b="<"/>', '<a/ >'])
    add('xml', F, fx(xml, '/*'))
  // namespaces: a default one, a declared prefix, an undeclared one, the functions over names
  for (const [xml, x] of [['<a xmlns="u"><b>1</b></a>', '//b'], ['<a xmlns="u"><b>1</b></a>', '//*[local-name()="b"]'], ['<a xmlns="u"><b>1</b></a>', 'namespace-uri(/*)'], ['<x:a xmlns:x="u">1</x:a>', '/*'], ['<x:a xmlns:x="u">1</x:a>', '//x:a'], ['<x:a>1</x:a>', '/*'], ['<a xmlns:x="u" x:k="1"/>', '/a/@*'], ['<a xmlns:x="u" x:k="1"/>', '/a/@k'], ['<a xmlns="u" k="1"/>', '/*/@k'], ['<x:a xmlns:x="u"><b>2</b></x:a>', '//b'], ['<a xml:lang="en"><b>1</b></a>', '//b[lang("en")]'], ['<a xmlns:x="u"/>', '/a/@*']])
    add('ns', F, fx(xml, x))
  // arguments: cells, numbers, booleans, arrays for either, many nodes
  add('args', F, `${F}(X!A1,X!A2)`, `INDEX(${F}(X!A1,"//b"),2)`, `${F}(X!A4,"/a")+1`, `${F}(X!A3,"/*")`, `${F}(TRUE,"/*")`, `${F}("<a>1</a>",TRUE)`, `${F}("<a>1</a>",5)`, `ROWS(${F}(X!A5,"//i"))`, `INDEX(${F}(X!A5,"//i"),300)`, `${F}(X!A5,"//i[last()]")`)
  add('args', F, `${F}({"<a>1</a>","<a>2</a>"},"/a")`, `INDEX(${F}({"<a>1</a>","<a>2</a>"},"/a"),1,2)`, `${F}("<a><b>1</b></a>",{"//b","/a"})`, `INDEX(${F}("<a><b>1</b><b>2</b></a>",{"//b","/a"}),1,2)`, `${F}(X!A1:A2,"//b")`, `${F}(1/0,"/a")`, `${F}("<a/>",1/0)`, `${F}(,"/a")`, `${F}("<a>1</a>",)`)
  return out
}

const PROBES = build()

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./filterxml.predicted.json', import.meta.url), 'utf8'))

/**
 * The predictions Excel for Windows did not confirm: a node's value is its own text, not its descendants' (an element
 * without text, a comment and an instruction are #VALUE!, `<a>x<b>1</b>y</a>` is `xy`), trimmed; TRUE and FALSE are
 * booleans; names are compared as written, so `//b` finds `b` under a default namespace and `//x:a` needs no declared
 * prefix; the namespace axis has nodes; and the parser had refused names with a colon.
 */
export const MISSED: readonly string[] = [
  'nodes#9', 'nodes#10', 'nodes#11', 'nodes#12', 'nodes#13', 'nodes#14', 'values#10', 'values#11', 'values#22',
  'values#23', 'values#24', 'values#26', 'values#29', 'values#31', 'values#32', 'xpath#11', 'xpath#12', 'xpath#14',
  'xpath#15', 'xpath#20', 'xpath#22', 'xpath#41', 'xpath#44', 'xpath#50', 'xpath-off#24', 'ns#1', 'ns#4', 'ns#5',
  'ns#7', 'ns#10', 'ns#11', 'args#14',
]
