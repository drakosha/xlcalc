/**
 * Data validation and outlines against Excel for Windows (`test/excel/win/valid/`, `bench/windows/valid.ts`): the Mac's
 * sets, with the differences both Excels share pinned in `VALID_PINS`.
 */
import { describe, expect, it } from 'vitest'
import { checkValidSets, VALID_PINS, validSets } from './validation-excel.js'

const SETS = validSets('win/valid/')

// COM in ru-RU shows a literal list with the locale's separator (`a;b;c` for the file's `"a,b,c"`), which reads both ways
const PINS = { ...VALID_PINS, loose: { 'ru-RU': 'lists' as const } }

describe('data validation and outlines against Excel for Windows', () => {
  it('every set of the Mac is answered', () => expect(SETS.map((s) => `${s.tag} ${s.name}`)).toEqual(validSets('valid/').map((s) => `${s.tag} ${s.name}`)))
})

checkValidSets(SETS, PINS, (name) => import(`../bench/coverage/valid/${name}.ts`))
