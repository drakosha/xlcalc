/**
 * Child process: computes the cases in formulajs from the number `start` and prints a JSON line per case.
 * The parent (`run.ts fj`) kills it when a case hangs and goes on from the next one.
 */
import { readFileSync } from 'node:fs'
import type { CaseSet } from './cases.js'
import { evalCase } from './formulajs.js'

const [file, start] = process.argv.slice(2)
const { cases } = JSON.parse(readFileSync(file, 'utf8')) as CaseSet
for (let i = Number(start); i < cases.length; i++) {
  process.stdout.write(JSON.stringify({ i, r: evalCase(cases[i]) }) + '\n')
}
