/** The current engine's file-width predictions and the object-model hypotheses, saved before the oracle. */
import { writeFileSync } from 'node:fs'
import { columnPoints } from '../../../src/numfmt/glyphs.js'
import { FILE_WIDTHS, SET_WIDTHS } from './cases.js'
const predictions = {
  Mac: { default: columnPoints(undefined), file: FILE_WIDTHS.map(columnPoints), set: SET_WIDTHS.map((c) => Math.round(c * 6 + 5)) },
  Windows: { default: 48, file: FILE_WIDTHS.map((w) => Math.trunc((256 * w + Math.trunc(128 / 7)) * 7 / 256) * 3 / 4), set: SET_WIDTHS.map((c) => Math.round(c * 7 + 5) * 3 / 4) },
}
writeFileSync(new URL('./width-held.predicted.json', import.meta.url), JSON.stringify(predictions, null, 2) + '\n')
