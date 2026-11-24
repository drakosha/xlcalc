import { writePrediction } from '../stored.js'
import { FORMAT_CONDITION_CASES, conditionFormatPrediction } from './number-format-conditions.js'

const entries = await Promise.all(FORMAT_CONDITION_CASES.map(async ({ id, code }) => [id, await conditionFormatPrediction(code)]))
writePrediction(import.meta.url, 'number-format-conditions', Object.fromEntries(entries))
