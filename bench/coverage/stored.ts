/**
 * Stored answers, as every step of both harnesses keeps them: read to go on from (`FCOV_FRESH=1` starts over), kept while
 * a case is the same, written in the order of the cases, so a run that adds or reruns cases changes the lines of those
 * cases only, however the run went. A set's module by its checked name, the run over the cases to do and its last line.
 * And the number format of a stored case or answer by either of its names.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import type { NumberFormat } from '../../src/contract.js'

/** A step's stored file to go on from; null without one, or with `FCOV_FRESH=1`, which runs every case again. */
export const loadStored = <T>(file: string): T | null => (existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as T) : null)

/** A file of answers as every step writes it: JSON indented by one space, a newline at the end; its folder made. */
export function writeAnswers(file: string, body: unknown): void {
  mkdirSync(dirname(file), { recursive: true })
  writeFileSync(file, JSON.stringify(body, null, 1) + '\n')
}

/** The seconds since `t0` to a tenth, as an answer file's `meta` gives its run's time. */
export const secondsSince = (t0: number): number => Math.round((Date.now() - t0) / 100) / 10

/** The seconds since `t0` as a run's log prints them. */
export const took = (t0: number): string => ((Date.now() - t0) / 1000).toFixed(1)

/** A case set's module, `bench/coverage/<kind>/<set>.ts`, by a name that is a file name and nothing more. */
export async function setModule<T>(kind: string, set: string | undefined): Promise<T> {
  if (!/^[a-z0-9-]+$/.test(set ?? '')) throw new Error(`set name: [a-z0-9-], not ${set}`)
  return (await import(`./${kind}/${set}.js`)) as T
}

/** The answers of `rec` in the order of `cases` (ids without an answer left out). */
export const inCaseOrder = <T>(cases: readonly { id: string }[], rec: Record<string, T>): Record<string, T> =>
  Object.fromEntries(cases.flatMap((c) => (c.id in rec ? [[c.id, rec[c.id]]] : [])))

/**
 * The stored answers a run keeps: those whose case is the same by `key` (its text unless given) in the stored file and
 * now. `key` gets the set the case belongs to, for a case whose book lies outside it.
 */
export function keptSame<C extends { id: string }, A, S extends { cases: C[] } = { cases: C[] }>(now: S, prev: (S & { answers: Record<string, A> }) | null, key: (c: C, set: S) => string = (c) => JSON.stringify(c)): Record<string, A> {
  const was = new Map(prev?.cases.map((c) => [c.id, key(c, prev)]) ?? [])
  const answers: Record<string, A> = {}
  for (const c of now.cases) if (prev?.answers[c.id] && was.get(c.id) === key(c, now)) answers[c.id] = prev.answers[c.id]
  return answers
}

/** A case's text without what the engine predicted for it: a prediction changed keeps Excel's answer. */
export const unpredicted = (c: object): string => JSON.stringify({ ...c, predicted: undefined })

/** The cases the variable `env` (`FSTRUCT_ONLY=id,id`) names, every case without it. */
export function chosen(env: string): (c: { id: string }) => boolean {
  const ids = process.env[env]?.split(',')
  return (c) => !ids || ids.includes(c.id)
}

/**
 * Runs the cases to do in turn, but those the variable `only` (`FSTRUCT_ONLY=id,id`) leaves out; `end` gets the number of
 * cases run however the run ends (the answers saved, its line printed).
 */
export async function runCases<C extends { id: string }>(todo: C[], only: string, run: (c: C) => Promise<void>, end: (run: number) => void): Promise<void> {
  let done = 0
  try {
    const take = chosen(only)
    for (const c of todo) {
      if (!take(c)) continue
      await run(c)
      done++
    }
  } finally {
    end(done)
  }
}

/** The line a run ends with: the cases it ran, how many of the set's are answered, and for a step that keeps refusals how many Excel did not load. */
export const answeredLine = (label: string, run: number, answers: object, cases: readonly unknown[], rejected?: object): string =>
  `${label}: ${run} cases run, ${Object.keys(answers).length} of ${cases.length} answered${rejected ? `, ${Object.keys(rejected).length} not loaded` : ''}`

/** A stored file of answers to cases sent as files: what was sent for each case (`partsHash`), the answers, the books Excel did not load. */
export interface StoredAnswers<C extends { id: string }, A> {
  cases: C[]
  sent: Record<string, string>
  answers: Record<string, A>
  rejected?: Record<string, string>
}

/** The stored answers a run keeps: a case's answer (its saved file there too) or refusal, while the case and its file are the same. */
export function keptAnswers<C extends { id: string }, A>(cases: C[], sent: Record<string, string>, prev: StoredAnswers<C, A> | null, saved: (id: string) => boolean): { answers: Record<string, A>; rejected: Record<string, string> } {
  const key = (c: C, hash: string | undefined) => JSON.stringify([c, hash])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.sent[c.id])]) ?? [])
  const answers: Record<string, A> = {}
  const rejected: Record<string, string> = {}
  for (const c of cases)
    if (same.get(c.id) === key(c, sent[c.id])) {
      if (prev!.answers[c.id] && saved(c.id)) answers[c.id] = prev!.answers[c.id]
      else if (prev!.rejected?.[c.id]) rejected[c.id] = prev!.rejected[c.id]
    }
  return { answers, rejected }
}

/** A case's answer with the file Excel saved, or why it has none. */
export type FileAnswer<A> = { answer: A; saved: Uint8Array } | { rejected: string }

/** The stored answers of a step whose cases go to Excel as files and come back saved (`fileAnswers`). */
export interface FileAnswers<C extends { id: string }, A> {
  cases: C[]
  answers: Record<string, A>
  rejected: Record<string, string>
  /** The cases with neither an answer nor a refusal kept. */
  todo: C[]
  /** Takes a case's answer (its saved file into the folder beside the answers' file) or refusal. */
  got(id: string, r: FileAnswer<A>): void
  /** Writes the answers' file: `meta`, the set's own fields, what was sent, the answers and the refusals, in the order of the cases. */
  save(meta: object, own: object): void
}

/** The answers of a step whose cases go to Excel as files (`sent` their `partsHash`), kept from `file` as `keptAnswers` keeps them, the saved files in the folder of its name. */
export function fileAnswers<C extends { id: string }, A>(file: string, cases: C[], sent: Record<string, string>): FileAnswers<C, A> {
  const dir = file.slice(0, -'.json'.length)
  const { answers, rejected } = keptAnswers<C, A>(cases, sent, loadStored(file), (id) => existsSync(join(dir, `${id}.xlsx`)))
  return {
    cases,
    answers,
    rejected,
    todo: cases.filter((c) => !answers[c.id] && !rejected[c.id]),
    got(id, r) {
      if ('rejected' in r) return void (rejected[id] = r.rejected)
      answers[id] = r.answer
      mkdirSync(dir, { recursive: true })
      writeFileSync(join(dir, `${id}.xlsx`), r.saved)
    },
    save(meta, own) {
      const refused = inCaseOrder(cases, rejected)
      writeAnswers(file, { meta, ...own, sent: inCaseOrder(cases, sent), answers: inCaseOrder(cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) })
    },
  }
}

/**
 * The number format of a stored case or answer: the books, cells and answers stored before the field was named
 * `numberFormat` name it `format`, and so do the case generators, whose cases are compared with their stored copies. The
 * engine takes both names and answers by the new one; the harness reads either.
 */
export const numberFormatOf = (x: { numberFormat?: NumberFormat; format?: NumberFormat } | null | undefined): NumberFormat | undefined => x?.numberFormat ?? x?.format

/**
 * Runs the cases to do of a step whose cases go to Excel as files, as `runCases` does: each answer or refusal goes into
 * `store` (a refusal named as it comes) and the answers are saved after each case and at the end.
 */
export function runFileCases<C extends { id: string }, A>(label: string, store: FileAnswers<C, A>, todo: C[], only: string, run: (c: C) => Promise<FileAnswer<A>>, save: () => void, t0: number): Promise<void> {
  return runCases(
    todo,
    only,
    async (c) => {
      const r = await run(c)
      if ('rejected' in r) console.log(`${c.id}: ${r.rejected}`)
      store.got(c.id, r)
      save()
    },
    (done) => {
      save()
      console.log(`${answeredLine(label, done, store.answers, store.cases, store.rejected)}, ${took(t0)} s`)
    },
  )
}
