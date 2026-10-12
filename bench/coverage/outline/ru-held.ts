/**
 * The commands in ru-RU, whose object model takes the same ranges and arguments: whether Excel for Windows in Russian
 * groups, hides and shows as in en-US. Held out, predicted before Excel ran it as `rules-held` was.
 */
import { collector, type OutlineCase } from './cases.js'
import { BOOKS } from './rules-held.js'

export { BOOKS }

export function buildOutlineCases(): OutlineCase[] {
  const { out, add } = collector()
  add('ru', 'nested', [{ hide: '8' }, { hide: '11' }, { show: '11' }])
  add('ru', 'nested', [{ levels: [1, 0] }, { levels: [2, 0] }])
  add('ru', 'empty', [{ group: '3:5' }, { ungroup: '4' }])
  add('ru', 'cols', [{ hide: 'K' }, { ungroup: 'A' }])
  return out
}
