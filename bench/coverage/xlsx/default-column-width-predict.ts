import { fileURLToPath } from 'node:url'
import { loadStored, writePrediction } from '../stored.js'
import { DEFAULT_WIDTH_CASES, defaultWidthPrediction } from './default-column-width.js'

const file = fileURLToPath(new URL('./default-column-width.predicted.json', import.meta.url))
const predictions = loadStored<Record<string, unknown>>(file) ?? {}
const added = DEFAULT_WIDTH_CASES.filter(({ id }) => !(id in predictions))
const entries = await Promise.all(added.map(async (c) => [c.id, await defaultWidthPrediction(c)]))
writePrediction(import.meta.url, 'default-column-width', { ...predictions, ...Object.fromEntries(entries) })
