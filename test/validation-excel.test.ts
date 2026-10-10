/**
 * Data validation and outlines against Excel 16 for Mac (`test/excel/valid/`, see `test/validation-excel.ts`): Excel read
 * in each book what the engine holds after the case (the book as written, with what Excel made, after the structure
 * operation), and the reader reads each file Excel saved as Excel read it. Known differences are pinned with their
 * reasons in `VALID_PINS` and, for what only Excel for Mac does, here.
 */
import { describe, expect, it } from 'vitest'
import { checkValidSets, VALID_PINS, validSets } from './validation-excel.js'

const SETS = validSets('valid/')

/**
 * In ru-RU Excel for Mac's AppleScript makes a list typed as `x;y;z` one item (Windows's COM splits it at the locale's
 * separator), and makes no rule of a formula with a Russian function name (`=ЕЧИСЛО(F2)`).
 */
const MAC_RU = ['S value C2: Excel fails, book meets', ...VALID_PINS.known['ru-RU/ru:typed'], ...['F2', 'F3'].map((a) => `S rule ${a}: Excel none, book {"type":"custom","allowBlank":true,"showPrompt":true,"showError":true,"errorStyle":"stop"}`)]

const PINS = {
  ...VALID_PINS,
  known: { ...VALID_PINS.known, 'ru-RU/ru:typed': MAC_RU },
  // the object model of Excel for Mac in ru-RU shows a rule's formulas now in Russian, now in English, cell by cell
  loose: { 'ru-RU': 'formulas' as const },
}

describe('data validation and outlines against Excel for Mac', () => {
  it('every set is answered', () => expect(SETS.map((s) => `${s.tag} ${s.name}`)).toEqual(['en-US entry-held-2', 'en-US entry-held', 'en-US files', 'en-US held-3', 'en-US lists', 'en-US made', 'en-US struct-held-2', 'en-US struct-held', 'ru-RU ru']))
})

checkValidSets(SETS, PINS, (name) => import(`../bench/coverage/valid/${name}.ts`))
