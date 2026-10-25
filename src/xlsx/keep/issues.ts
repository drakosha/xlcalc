/** The issues of writing over a file: those of the parts written from the model alone are taken (`copyTo`). */
import { Issues } from '../issues.js'
import type { Where, XlsxMessage } from '../issues.js'
import type { XlsxIssueCode } from '../types.js'

export class KeptIssues extends Issues {
  /** Each issue gathered, cell by cell, added to `to` where `keep` takes it. */
  copyTo(to: Issues, keep: (code: XlsxIssueCode, message: XlsxMessage, where: Where) => boolean): void {
    for (const g of this.byKey.values())
      for (const cell of g.cells ?? [undefined]) {
        const where = cell === undefined ? g.where : { ...g.where, cell }
        if (keep(g.code, g.message, where)) to.add(g.code, g.message, where)
      }
  }
}
