/**
 * Issues gathered while reading or writing: one per code, place and message, with the cells it is about in order. A
 * message is kept as its key and parameters and written in the locale of the call when the issues are listed.
 */
import type { Message, XlsxMessageParams } from '../locale/messages.js'
import type { Locale } from '../locale/types.js'
import { issueText } from './messages.js'
import type { XlsxIssue, XlsxIssueCode } from './types.js'

/** A message of an issue. */
export type XlsxMessage = Message<keyof XlsxMessageParams>

export interface Where {
  sheet?: string
  cell?: string
  name?: string
}

interface Gathered {
  code: XlsxIssueCode
  message: XlsxMessage
  where: Where
  cells?: string[]
}

export class Issues {
  private readonly byKey = new Map<string, Gathered>()

  add(code: XlsxIssueCode, message: XlsxMessage, where: Where = {}): void {
    const key = `${code}\n${where.sheet ?? ''}\n${where.name ?? ''}\n${JSON.stringify(message)}`
    let issue = this.byKey.get(key)
    if (!issue) this.byKey.set(key, (issue = { code, message, where }))
    if (where.cell !== undefined) (issue.cells ??= []).push(where.cell)
  }

  /** Each issue gathered, cell by cell, added to `to` where `keep` takes it. */
  copyTo(to: Issues, keep: (code: XlsxIssueCode, message: XlsxMessage, where: Where) => boolean): void {
    for (const g of this.byKey.values())
      for (const cell of g.cells ?? [undefined]) {
        const where = cell === undefined ? g.where : { ...g.where, cell }
        if (keep(g.code, g.message, where)) to.add(g.code, g.message, where)
      }
  }

  /** The issues, their messages in `locale` (en-US without one). */
  list(locale?: Locale): XlsxIssue[] {
    return [...this.byKey.values()].map(({ code, message, where, cells }) => ({
      code,
      message: issueText(locale, message),
      ...(where.sheet !== undefined ? { sheet: where.sheet } : {}),
      ...(where.name !== undefined ? { name: where.name } : {}),
      ...(cells ? { cells } : {}),
    }))
  }
}
