/**
 * The three ways a session runs, behind the asynchronous API so that one test covers them: a synchronous session, a
 * session computed in portions in the same thread and one in a thread of the pool.
 */
import { beforeAll } from 'vitest'
import { createInlineCalc, createSession, type AsyncSession, type BookInput, type Calc, type Session, type SessionOptions } from '../src/index.js'
import { pool } from './fixtures/pool.js'

export const KINDS = ['a synchronous session', 'the same thread', 'a thread'] as const
export type Kind = (typeof KINDS)[number]

/** A synchronous session behind the asynchronous API: each of its methods as a promise, its clock, history, revision and `subscribe` as they are. */
export function wrapped(s: Session): AsyncSession {
  const out: Record<string, unknown> = {}
  for (const k of Object.keys(s) as (keyof Session)[]) {
    if (k === 'clock' || k === 'history' || k === 'revision') Object.defineProperty(out, k, { get: () => s[k] })
    // subscribing is synchronous in every kind of session
    else if (k === 'subscribe') out[k] = s[k]
    else out[k] = async (...a: unknown[]) => (s[k] as (...a: unknown[]) => unknown)(...a)
  }
  return out as unknown as AsyncSession
}

const calcs: Calc[] = []

/** Closes the calculators the openers made (a file's `afterAll`). */
export const closeCalcs = () => Promise.all(calcs.splice(0).map((c) => c.close()))

/** A pool kept to be closed with the others. */
export async function keptPool(): Promise<Calc> {
  const p = await pool({ threads: 1 })
  calcs.push(p)
  return p
}

/** Opens sessions of one kind. */
export async function opener(kind: Kind): Promise<(book: BookInput, opts: SessionOptions) => Promise<AsyncSession>> {
  if (kind === 'a synchronous session') return async (book, opts) => wrapped(createSession(book, opts))
  const calc = kind === 'the same thread' ? createInlineCalc({ sliceMs: 2 }) : await keptPool()
  if (kind === 'the same thread') calcs.push(calc)
  return (book, opts) => calc.createSession(book, opts)
}

/** Opens sessions of one kind in a `describe`'s tests: the opener is made before them (`beforeAll`). */
export function openerFor(kind: Kind): (book: BookInput, opts: SessionOptions) => Promise<AsyncSession> {
  let open: Awaited<ReturnType<typeof opener>>
  beforeAll(async () => {
    open = await opener(kind)
  })
  return (book, opts) => open(book, opts)
}
