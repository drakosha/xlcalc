/**
 * What Excel for Windows does with a LAMBDA's optional parameter and a text literal over 255 characters when they are typed, and how it
 * saves them: `enter` runs on Windows alone. The spec's grammar has `_xlop.` for a LAMBDA's parameter and `_xlfn.LONGTEXT` of string
 * constants ([MS-XLSX] 2.2.2) but says nothing of what they mean: that `_xlop.` is `[y]`, and that a long literal is saved as LONGTEXT,
 * are guesses this set tests. The guesses are `predicted`, committed before Excel ran the set; the saved files are the answers.
 */
import type { CachedValue, LinksCase, LinksCases } from './cases.js'

const VALUE = { error: 'VALUE' } as const

function typed(id: string, about: string, formula: string, predicted?: CachedValue): LinksCase {
  return {
    id,
    about,
    sheets: [{ name: 'S', cells: { B1: { value: 1 } } }],
    linked: {},
    links: [],
    read: [{ sheet: 'S', range: 'A2' }],
    ops: [{ op: 'enter', sheet: 'S', cell: 'A2', formula }],
    ...(predicted === undefined ? {} : { predicted: { 'S!A2': predicted } }),
  }
}

const OPTIONAL = 'LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))'

export function buildLinksCases(): LinksCases {
  return {
    cases: [
      typed('optional-one', 'the optional parameter left out', `=${OPTIONAL}(1)`, 1),
      typed('optional-two', 'the optional parameter given', `=${OPTIONAL}(1,2)`, 3),
      typed('optional-three', 'one argument too many', `=${OPTIONAL}(1,2,3)`, VALUE),
      typed('optional-empty', 'the optional parameter given empty', `=${OPTIONAL}(1,)`, 1),
      typed('required-after', 'a required parameter after an optional one: Excel refuses the typing, the cell stays empty', '=LAMBDA([x],y,IF(ISOMITTED(x),y,x))(1,2)'),
      typed('literal-255', 'a literal of 255 characters', `=LEN("${'a'.repeat(255)}")`, 255),
      typed('literal-256', 'a literal of 256 characters', `=LEN("${'a'.repeat(256)}")`, 256),
      typed('literal-4095', 'a literal of 4,095 characters', `=LEN("${'a'.repeat(4095)}")`, 4095),
      typed('literal-4096', 'a literal of 4,096 characters', `=LEN("${'a'.repeat(4096)}")`),
    ],
  }
}
