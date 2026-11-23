import { writePrediction } from '../stored.js'
import { DEFAULT_WIDTH_CASES, defaultWidthPrediction } from './default-column-width.js'

const entries = await Promise.all(DEFAULT_WIDTH_CASES.map(async (c) => [c.id, await defaultWidthPrediction(c)]))
writePrediction(import.meta.url, 'default-column-width', Object.fromEntries(entries))
