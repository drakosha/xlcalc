/**
 * `numberFormatCategory`'s names for the cells of a ui set's ribbon cases as the predictions of the Windows step
 * `ribbon`, written before Excel runs it: `npx tsx bench/windows/ribbon-predict.ts <set>` writes
 * `ribbon/<set>.predicted.json`, each case's cells by name in each of the package's locales.
 */
import { fileURLToPath } from 'node:url'
import { TAGS } from '../coverage/locales.js'
import { writeAnswers } from '../coverage/stored.js'
import { ribbonOf } from '../coverage/ui/cases.js'
import { ribbonCases } from './ribbon.js'

const set = process.argv[2]
const cases = await ribbonCases(set)
const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(cases.map((c) => [c.id, ribbonOf(c, tag)]))]))
writeAnswers(fileURLToPath(new URL(`./ribbon/${set}.predicted.json`, import.meta.url)), { now: new Date().toISOString().slice(0, 10), locales })
