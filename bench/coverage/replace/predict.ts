/**
 * A replace set's predictions, written before Excel runs the set: `npx tsx bench/coverage/replace/predict.ts <set> [grid]`
 * writes `replace/<set>.predicted.json`, each case's cells it changed (their value or formula and format, as `heldOf`
 * gives them, null for a cell left empty) in each of the package's locales, and whether the model refused it. By
 * default they are the engine's of the day (`replaceInBook`: `replaceInCells`); with `grid`, `modelReplace`, a grid's
 * Replace All with the matches `?` reads as Find does (for `rules-held`).
 */
import { TAGS } from '../locales.js'
import { setModule, writePrediction } from '../stored.js'
import { changedCells, modelReplace, replaceInBook, type ReplaceSet } from './cases.js'

const [set, model] = process.argv.slice(2)
const { BOOKS, buildReplaceCases } = await setModule<ReplaceSet>('replace', set)
const cases = buildReplaceCases()
const locales = Object.fromEntries(
  TAGS.map((tag) => [
    tag,
    Object.fromEntries(
      cases.map((c) => {
        const m = model === 'grid' ? modelReplace(BOOKS[c.book], c.query, tag) : { cells: replaceInBook(BOOKS[c.book], c.query, tag), refused: undefined }
        return [c.id, { ...(m.refused ? { refused: m.refused } : {}), changed: changedCells(BOOKS[c.book], m.cells, tag) }]
      }),
    ),
  ]),
)
writePrediction(import.meta.url, `${set}`, { model: model === 'grid' ? 'modelReplace: a grid Replace All, Find in formulas and setCellTexts' : 'replaceInCells', locales })
