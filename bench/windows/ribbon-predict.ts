/**
 * `numberFormatCategory`'s names for the codes of the Windows step `ribbon` as its predictions, written before Excel
 * runs the step: `npx tsx bench/windows/ribbon-predict.ts` writes `ribbon.predicted.json`, the name for each code of
 * `RIBBON_CODES` in each of the package's locales.
 */
import { fileURLToPath } from 'node:url'
import { numberFormatCategory } from '../../src/index.js'
import { TAGS } from '../coverage/locales.js'
import { writeAnswers } from '../coverage/stored.js'
import { RIBBON_CODES } from './ribbon.js'

const locales = Object.fromEntries(TAGS.map((tag) => [tag, RIBBON_CODES.map((c) => numberFormatCategory(c, tag).name)]))
writeAnswers(fileURLToPath(new URL('./ribbon.predicted.json', import.meta.url)), { now: new Date().toISOString().slice(0, 10), codes: RIBBON_CODES, locales })
