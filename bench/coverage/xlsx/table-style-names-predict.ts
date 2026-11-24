import { writePrediction } from '../stored.js'
import { TABLE_STYLE_CASES, tableStylePrediction } from './table-style-names.js'

const answers = await Promise.all(TABLE_STYLE_CASES.map(async (c) => [c.id, await tableStylePrediction(c)]))
writePrediction(import.meta.url, 'table-style-names', Object.fromEntries(answers))
