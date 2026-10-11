/**
 * The old names of the API's methods, deprecated, each the same function as its new name until hosts move to the new
 * names: a session's (`book` is `getBook`, `restructure` is `editStructure` and the rest) and a calculator's
 * (`openSession` is `createSession`).
 */
import type { Calc } from './contract.js'

/** A session's methods by their old names, and the new name of each. */
const SESSION = { restructure: 'editStructure', book: 'getBook', layout: 'getLayout', extent: 'getUsedRange', cells: 'getCells', validation: 'getValidation', result: 'getResult', hash: 'getHash' } as const

type SessionOldName = keyof typeof SESSION

/** A session (synchronous or asynchronous) with its methods' old names added, each the method of its new name. */
export function withOldNames<T extends Record<SessionOldName, unknown>>(session: Omit<T, SessionOldName>): T {
  const s = session as Record<string, unknown>
  for (const [old, name] of Object.entries(SESSION)) s[old] = s[name]
  return session as T
}

/** A calculator with `openSession`, the old name of `createSession`. */
export const calcWithOldNames = (calc: Omit<Calc, 'openSession'>): Calc => Object.assign(calc, { openSession: calc.createSession })
