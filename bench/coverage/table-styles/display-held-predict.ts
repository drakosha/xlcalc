import { createSession, registerPivot } from '../../../src/index.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { readXlsx } from '../../../src/xlsx.js'
import { writePrediction } from '../stored.js'
import { DISPLAY_CASES, displayFile } from './display-held.js'

registerPivot(pivotEngine)
const predictions = await Promise.all(DISPLAY_CASES.map(async ({ id, sheet, range }) => {
  const { book, cached } = await readXlsx(await displayFile(id))
  const session = createSession(book, { cached })
  try {
    const window = session.getViewport(sheet, range)
    const cells = Object.fromEntries(Object.entries(window.cells).map(([address, cell]) => [
      address, { style: cell.effectiveStyle === undefined ? null : window.styles[cell.effectiveStyle], appearance: cell.appearance ?? null },
    ]))
    return [id, cells]
  } finally {
    session.close()
  }
}))
writePrediction(import.meta.url, 'display-held', Object.fromEntries(predictions))
