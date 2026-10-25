/**
 * Which sheet of the book each sheet of the file became, read from the carriers (`carriers.ts`), and the changes
 * `writeKept` refuses: those that move a reference of a part the model does not hold (a carrier that moved), that
 * change a sheet whose parts hold references no carrier reads, that rename or delete what a part names (a pivot
 * chart's sheet, a pivot cache's table or name), that delete a sheet with parts of its own the model does not have, or
 * that change the links to other workbooks. The file would point at the wrong cells: no file is written then.
 */
import type { BookInput } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { keyOf } from '../../style-table.js'
import { sheetKey } from '../../syntax/print.js'
import { compared, lostSheet, sheetOfCarrier, type Carriers } from './carriers.js'
import type { Source } from './source.js'

/** `writeKept`'s refusal: the change would leave a part of the file the model does not hold pointing at the wrong cells. */
export class KeptRefusal extends Error {
  override readonly name = 'KeptRefusal'
  constructor(why: string) {
    super(`xlsx/keep: ${why}`)
  }
}

export interface Identity {
  /** By sheet of the book written: the file's sheet it is (an index of the book read), none for a new one. */
  source: (number | undefined)[]
  /** By sheet of the book read: the sheet of the book written it became; none for one deleted. */
  target: Map<number, number>
  /** By sheet of the book read: rows or columns were inserted or deleted on it. */
  lines: boolean[]
}

/** The relationships a sheet deleted may have: parts of its own the model has, or that hold no reference. */
const DELETABLE = new Set(['table', 'comments', 'vmlDrawing', 'printerSettings', 'hyperlink', 'threadedComment'])

const OPEN_KEPT = 'open the file with openKept and keep the hidden names it adds'

export function refusals(src: Source, book0: BookInput, book1: BookInput, c: Carriers): Identity {
  const names = new Map((book1.names ?? []).filter((n) => n.scope === undefined).map((n) => [caseKey(n.name), n]))
  const index = new Map(book1.sheets.map((s, j) => [sheetKey(s.name), j]))
  const out: Identity = { source: book1.sheets.map(() => undefined), target: new Map(), lines: [] }
  for (const r of c.refs) {
    const n = names.get(caseKey(r.name))
    if (!n) throw new KeptRefusal(`the workbook has no name ${r.name}, which holds a reference of ${r.what}: ${OPEN_KEPT}`)
    // a part on a sheet deleted goes with it (as far as `DELETABLE` lets it)
    if (r.kind === 'place' && lostSheet(n.ref)) continue
    if (compared(n.ref, r.kind) !== compared(r.ref, r.kind))
      throw new KeptRefusal(`${r.what} refers to ${r.ref}, which the change made ${n.ref}; writeKept does not move the references of parts the workbook does not hold, so the file would point at other cells`)
  }
  for (const [k, carrier] of c.sheets.entries()) {
    const was = book0.sheets[k].name
    const n = names.get(caseKey(carrier.name))
    const at = n && sheetOfCarrier(n.ref)
    const j = n ? at && index.get(sheetKey(at.sheet)) : index.get(sheetKey(was))
    if (j !== undefined && out.source[j] === undefined) {
      out.source[j] = k
      out.target.set(k, j)
    }
    out.lines[k] = !!at?.moved
    const renamed = j === undefined || book1.sheets[j].name !== was
    const changed = renamed || out.lines[k]
    const unread = [...c.unparsed[k], ...c.anywhere]
    if (!n && (unread.length || c.named[k].length)) throw new KeptRefusal(`the workbook has no name ${carrier.name}, which tells what became of the sheet '${was}': ${OPEN_KEPT}`)
    if (changed && unread.length) throw new KeptRefusal(`the sheet '${was}' had rows or columns inserted or deleted, or was renamed or deleted, and ${unread[0]} holds references writeKept does not read`)
    if (renamed && c.named[k].length) throw new KeptRefusal(`the sheet '${was}' was renamed or deleted, and ${c.named[k][0]} names it`)
    if (j === undefined) {
      const own = src.sheets[k].rels.find((r) => !DELETABLE.has(r.kind))
      if (own) throw new KeptRefusal(`the sheet '${was}' was deleted, and it has a part the workbook does not hold (${own.kind} ${own.target})`)
    }
  }
  const tables = new Set(book1.sheets.flatMap((s) => (s.tables ?? []).map((t) => caseKey(t.name))))
  for (const b of c.byName)
    if (!tables.has(caseKey(b.name)) && !(book1.names ?? []).some((n) => caseKey(n.name) === caseKey(b.name)))
      throw new KeptRefusal(`${b.what} is the table or name ${b.name}, which the workbook no longer has`)
  if (keyOf(book0.externalBooks ?? []) !== keyOf(book1.externalBooks ?? [])) throw new KeptRefusal('the links to other workbooks changed, which writeKept does not write')
  return out
}
