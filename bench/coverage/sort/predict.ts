/**
 * `npx tsx bench/coverage/sort/predict.ts <set>`: the order Microsoft's words on the Sort command give, into
 * `sort/<set>.predicted.json`, committed before Excel runs the set. Numbers, text in any case, FALSE then TRUE, errors
 * all equal, then blanks; descending reverses it but for blanks, last either way; equal cells keep their order.
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import type { SortSet } from './cases.js'

const name = process.argv[2]
const { set } = await setModule<{ set: SortSet }>('sort', name)

/** A cell's place among the types and its key within its type. */
function keyOf(t: string): [number, number | string] {
  if (t === '') return [4, 0]
  if (t.startsWith('=')) return [3, 0]
  if (t === 'FALSE' || t === 'TRUE') return [2, t === 'TRUE' ? 1 : 0]
  return /^-?\d+(\.\d+)?$/.test(t) ? [0, Number(t)] : [1, t]
}

function order(a: string, b: string, desc: boolean): number {
  const [ra, ka] = keyOf(a)
  const [rb, kb] = keyOf(b)
  if (ra === 4 || rb === 4) return ra - rb
  const c = ra !== rb ? ra - rb : typeof ka === 'string' ? ka.localeCompare(kb as string, 'en-US', { sensitivity: 'base' }) : (ka as number) - (kb as number)
  return desc ? -c : c
}

const answers = Object.fromEntries(set.cases.map((c) => [c.id, [...c.cells].sort((a, b) => order(a, b, !!c.descending))]))
writeAnswers(fileURLToPath(new URL(`./${name}.predicted.json`, import.meta.url)), { model: "Microsoft's Sort order, errors equal", answers })
