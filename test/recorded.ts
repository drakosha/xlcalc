/** Answers recorded as labelled JSON (`read-answers.ts`, `edit-answers.ts`), compared with those given now in every kind of session. */
import { readFileSync } from 'node:fs'
import { afterAll, describe, expect } from 'vitest'
import type { Execution } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

/** Answers in order: each a label and the answer's JSON. */
export type Labelled = [label: string, json: string][]

/** The recorded answers of a fixture file. */
export const readRecorded = <T>(file: string): T => JSON.parse(readFileSync(file, 'utf8')) as T

/** The answers given now are the recorded ones: the same labels in order, then each answer under its label. */
export function expectRecorded(got: Labelled, want: Labelled): void {
  expect(got.map(([label]) => label)).toEqual(want.map(([label]) => label))
  for (let i = 0; i < want.length; i++) expect(got[i], want[i][0]).toEqual(want[i])
}

/** Tests in a synchronous session, in the same thread and in a thread, with closures and eval: `body` gets the execution and an opener of the kind. */
export function inEverySession(title: string, body: (execution: Execution, open: ReturnType<typeof openerFor>) => void): void {
  afterAll(closeCalcs)
  describe.each(['closures', 'eval'] as Execution[])(`${title}, execution %s`, (execution) => {
    describe.each(KINDS)('%s', (kind) => body(execution, openerFor(kind)))
  })
}
