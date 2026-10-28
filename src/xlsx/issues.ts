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
  /** A user-defined function the issue is about (`XlsxIssue.fn`). */
  fn?: string
}

export interface Gathered {
  code: XlsxIssueCode
  message: XlsxMessage
  where: Where
  cells?: string[]
}

/** Two places of one issue: the same sheet, name and function, a missing one as an empty one. */
const same = (a: Where, b: Where) => (a.sheet ?? '') === (b.sheet ?? '') && (a.name ?? '') === (b.name ?? '') && (a.fn ?? '') === (b.fn ?? '')

export class Issues {
  protected readonly byKey = new Map<string, Gathered>()
  /** The issues a message object was added to, by its code and place: a caller that adds one message for many cells builds no key for each. */
  private readonly byMessage = new WeakMap<XlsxMessage, Gathered[]>()

  add(code: XlsxIssueCode, message: XlsxMessage, where: Where = {}): void {
    let seen = this.byMessage.get(message)
    let issue: Gathered | undefined
    for (const g of seen ?? [])
      if (g.code === code && same(g.where, where)) {
        issue = g
        break
      }
    if (!issue) {
      const key = `${code}\n${where.sheet ?? ''}\n${where.name ?? ''}\n${where.fn ?? ''}\n${JSON.stringify(message)}`
      issue = this.byKey.get(key)
      if (!issue) this.byKey.set(key, (issue = { code, message, where }))
      if (!seen) this.byMessage.set(message, (seen = []))
      seen.push(issue)
    }
    if (where.cell !== undefined) (issue.cells ??= []).push(where.cell)
  }

  /** The issues, their messages in `locale` (en-US without one). */
  list(locale?: Locale): XlsxIssue[] {
    return [...this.byKey.values()].map(({ code, message, where, cells }) => ({
      code,
      message: issueText(locale, message),
      ...(where.sheet !== undefined ? { sheet: where.sheet } : {}),
      ...(where.name !== undefined ? { name: where.name } : {}),
      ...(where.fn !== undefined ? { fn: where.fn } : {}),
      ...(cells ? { cells } : {}),
    }))
  }
}
