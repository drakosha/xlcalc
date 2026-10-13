/**
 * Cases of `rules-held` in ru-RU, whose object model takes the same cells: whether Excel for Windows in Russian goes as
 * in en-US over blocks, an empty text from a formula, a hidden row and a merged area. Held out, predicted before Excel
 * ran it as `rules-held` was.
 */
import { collector, type EndCase } from './cases.js'
import { BOOKS } from './rules-held.js'

export { BOOKS }

export function buildEndCases(): EndCase[] {
  const { out, add } = collector()
  add('ru', 'blocks', 'A1', 'down')
  add('ru', 'kinds', 'B1', 'down')
  add('ru', 'hidden', 'A1', 'down')
  add('ru', 'merged', 'A2', 'right')
  return out
}
