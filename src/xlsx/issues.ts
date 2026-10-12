/**
 * Issues gathered while reading or writing: one per code, place and message, with the cells it is about in order. A
 * message is kept as its key and parameters and written in the locale of the call when the issues are listed.
 */
import { checkLocaleArg } from '../locale/load.js'
import type { Message, XlsxMessageParams } from '../locale/messages.js'
import type { Locale } from '../locale/types.js'
import { issueText } from './messages.js'
import type { XlsxIssue, XlsxIssueCode } from './types.js'

/** A message of an issue. */
export type XlsxMessage = Message<keyof XlsxMessageParams>

/** `opts.locale` of `readXlsx` and `writeXlsx`, the language of their issues: a loaded locale's tag or a locale object. */
export const issueLocale = (o: Record<string, unknown>): Locale | undefined => (o.locale === undefined ? undefined : checkLocaleArg(o.locale, 'opts.locale'))

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
