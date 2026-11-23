import { writePrediction } from '../stored.js'
import { EMPTY_NAME_CASES, emptyNamePrediction } from './empty-defined-names.js'

const entries = await Promise.all(EMPTY_NAME_CASES.map(async (c) => [c.id, await emptyNamePrediction(c)]))
writePrediction(import.meta.url, 'empty-defined-names', Object.fromEntries(entries))
