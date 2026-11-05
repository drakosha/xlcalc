/**
 * A LAMBDA's optional parameter (`_xlop.`) and `_xlfn.LONGTEXT` in a file, each case a file of its own, as `links/no-prefix.ts`: a refused
 * file costs only it. `longtext-typed` showed `[y]` is saved as `_xlop.y`, with `_xlpm.y` in the body, and that typing a literal over 255
 * characters is refused, so nothing there says when Excel writes LONGTEXT. `predicted` and `REFUSED` were committed before Excel ran the
 * set: the grammar of [MS-XLSX] 2.2.2 (2 to 17 string constants, `_xlop.` on a LAMBDA alone) with the meaning `longtext-typed` measured.
 */
import type { CachedValue, LinksCase, LinksCases } from './cases.js'

const one = (id: string, about: string, formula: string, predicted?: CachedValue): LinksCase => ({
  id,
  about,
  sheets: [{ name: 'S', cells: { B1: { value: 1 } } }],
  linked: { S: { A1: formula } },
  links: [],
  read: [{ sheet: 'S', range: 'A1' }],
  ...(predicted === undefined ? {} : { predicted: { 'S!A1': predicted } }),
})

const OPTIONAL = '_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y))'
const text = (n: number, length: number) => Array.from({ length: n }, () => `"${'a'.repeat(length)}"`).join(',')

/** The cases predicted to make Excel refuse the file. */
export const REFUSED = ['xlop-after', 'xlop-let', 'xlop-in-body', 'lt-one', 'lt-eighteen', 'lt-cell', 'lt-number', 'literal-300']

export function buildLinksCases(): LinksCases {
  return {
    cases: [
      one('xlop-one', 'the optional parameter left out', `${OPTIONAL}(1)`, 1),
      one('xlop-two', 'the optional parameter given', `${OPTIONAL}(1,2)`, 3),
      one('xlop-three', 'one argument too many', `${OPTIONAL}(1,2,3)`, { error: 'VALUE' }),
      one('xlop-few', 'a required parameter left out', `${OPTIONAL}()`, { error: 'VALUE' }),
      one('xlop-blank', 'an optional parameter left out reads as a blank', '_xlfn.LAMBDA(_xlop.y,_xlpm.y+1)()', 1),
      one('xlop-both', 'every parameter optional, none given', '_xlfn.LAMBDA(_xlop.a,_xlop.b,IF(_xlfn.ISOMITTED(_xlpm.a),1,0)+IF(_xlfn.ISOMITTED(_xlpm.b),10,0))()', 11),
      one('xlop-after', 'a required parameter after an optional one', '_xlfn.LAMBDA(_xlop.x,_xlpm.y,_xlpm.y)(1,2)'),
      one('xlop-let', 'on a LET name, which the grammar does not take', '_xlfn.LET(_xlop.a,1,_xlpm.a)'),
      one('xlop-in-body', 'the prefix on a use of the parameter', '_xlfn.LAMBDA(_xlpm.x,_xlop.y,_xlpm.x+_xlop.y)(1,2)'),
      one('lt-two', 'two constants', '_xlfn.LONGTEXT("aaa","bbb")', 'aaabbb'),
      one('lt-len', 'inside LEN', 'LEN(_xlfn.LONGTEXT("aaa","bbb"))', 6),
      one('lt-empty', 'empty constants', 'LEN(_xlfn.LONGTEXT("",""))', 0),
      one('lt-quote', 'a doubled quote in a constant', '_xlfn.LONGTEXT("a""","b")', 'a"b'),
      one('lt-one', 'one constant, under the grammar', '_xlfn.LONGTEXT("aaa")'),
      one('lt-seventeen', '17 constants of 255: the most the grammar takes, past its 4,095 of SHOULD NOT', `LEN(_xlfn.LONGTEXT(${text(17, 255)}))`, 4335),
      one('lt-eighteen', '18 constants, over the grammar', `LEN(_xlfn.LONGTEXT(${text(18, 1)}))`),
      one('lt-long', 'two constants of 255', `LEN(_xlfn.LONGTEXT(${text(2, 255)}))`, 510),
      one('lt-cell', 'a reference for a constant', '_xlfn.LONGTEXT(B1,"a")'),
      one('lt-number', 'numbers for constants', '_xlfn.LONGTEXT(1,2)'),
      one('lt-bare', 'without its prefix', 'LONGTEXT("a","b")', { error: 'NAME' }),
      one('literal-255', 'a literal of 255 characters', `LEN("${'a'.repeat(255)}")`, 255),
      one('literal-300', 'a literal of 300 characters', `LEN("${'a'.repeat(300)}")`),
    ],
  }
}
