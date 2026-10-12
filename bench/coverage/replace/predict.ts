/**
 * A replace set's predictions, written before Excel runs the set: `npx tsx bench/coverage/replace/predict.ts <set>`
 * writes `replace/<set>.predicted.json`, each case's cells it changed (their value or formula and format, as `heldOf`
 * gives them, null for a cell left empty) in each of the package's locales, and whether the model refused it. The model is `modelReplace`: a grid's
 * Replace All, with the matches `?` reads as Find does.
 */
import { fileURLToPath } from 'node:url'
import { TAGS } from '../locales.js'
import { setModule, writeAnswers } from '../stored.js'
import { changedCells, modelReplace, type ReplaceSet } from './cases.js'

const [set] = process.argv.slice(2)
const { BOOKS, buildReplaceCases } = await setModule<ReplaceSet>('replace', set)
const cases = buildReplaceCases()
const locales = Object.fromEntries(
  TAGS.map((tag) => [
    tag,
    Object.fromEntries(
      cases.map((c) => {
        const m = modelReplace(BOOKS[c.book], c.query, tag)
        return [c.id, { ...(m.refused ? { refused: m.refused } : {}), changed: changedCells(BOOKS[c.book], m.cells, tag) }]
      }),
    ),
  ]),
)
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), { model: 'modelReplace: a grid Replace All, Find in formulas and setCellTexts', locales })
