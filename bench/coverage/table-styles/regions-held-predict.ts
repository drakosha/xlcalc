import { createSession } from '../../../src/index.js'
import '../../../src/locales/ru-RU.js'
import { registerTableStyles } from '../../../src/table-styles.js'
import { readXlsx } from '../../../src/xlsx.js'
import { writePrediction } from '../stored.js'
import { REGION_CASES, regionFile } from './regions-held.js'

registerTableStyles()
const predictions = await Promise.all(REGION_CASES.map(async ({ id, sheet, range }) => {
  const { book, cached } = await readXlsx(await regionFile(id))
  const session = createSession(book, { cached })
  try {
    const view = session.getViewport(sheet, range)
    const styles = Object.fromEntries(Object.entries(view.cells).flatMap(([address, { resolvedStyle }]) =>
      resolvedStyle === undefined ? [] : [[address, view.resolvedStyles![resolvedStyle]]]))
    return [id, styles]
  } finally {
    session.close()
  }
}))
writePrediction(import.meta.url, 'regions-held', Object.fromEntries(predictions))
