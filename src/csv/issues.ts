/** The issues of a CSV file read or written: one per message, with the cells it is about in order, its text in the locale of the call. */
import type { Locale } from '../locale/types.js'
import { csvText, type CsvMessage } from './messages.js'
import type { CsvIssue, CsvIssueCode } from './types.js'

export class CsvIssues {
  private readonly byKey = new Map<string, { code: CsvIssueCode; message: CsvMessage; cells?: string[] }>()

  constructor(private readonly sheet: string) {}

  add(code: CsvIssueCode, message: CsvMessage, cell?: string): void {
    const key = JSON.stringify(message)
    let issue = this.byKey.get(key)
    if (!issue) this.byKey.set(key, (issue = { code, message }))
    if (cell !== undefined) (issue.cells ??= []).push(cell)
  }

  list(locale: string | Locale | undefined): CsvIssue[] {
    return [...this.byKey.values()].map(({ code, message, cells }) => ({ code, message: csvText(locale, message), sheet: this.sheet, ...(cells ? { cells } : {}) }))
  }
}
