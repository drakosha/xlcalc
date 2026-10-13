/** Save the draft engine's clipboard predictions before running Excel. */
import { writeFileSync } from 'node:fs'
import { clipboardCases, predictedText } from './cases.js'
const answers = Object.fromEntries(['en-US', 'ru-RU'].map((tag) => [tag, Object.fromEntries(clipboardCases(tag).map((c) => [c.id, predictedText(c.rows)]))]))
writeFileSync(new URL('./copy-held.predicted.json', import.meta.url), JSON.stringify(answers, null, 2) + '\n')
