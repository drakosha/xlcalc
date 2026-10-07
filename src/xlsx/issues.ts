/** Issues gathered while reading or writing: one per code, place and message, with the cells it is about in order. */
import type { XlsxIssue, XlsxIssueCode } from './types.js'

export interface Where {
  sheet?: string
  cell?: string
  name?: string
}

export class Issues {
  private readonly byKey = new Map<string, XlsxIssue>()

  add(code: XlsxIssueCode, message: string, where: Where = {}): void {
    const key = `${code}\n${where.sheet ?? ''}\n${where.name ?? ''}\n${message}`
    let issue = this.byKey.get(key)
    if (!issue) {
      issue = { code, message, ...(where.sheet !== undefined ? { sheet: where.sheet } : {}), ...(where.name !== undefined ? { name: where.name } : {}) }
      this.byKey.set(key, issue)
    }
    if (where.cell !== undefined) (issue.cells ??= []).push(where.cell)
  }

  list(): XlsxIssue[] {
    return [...this.byKey.values()]
  }
}
