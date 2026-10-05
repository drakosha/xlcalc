/**
 * Дочерний процесс: считает случаи formulajs с номера `start` и печатает строку JSON на случай.
 * Родитель (`run.ts fj`) убивает процесс, если случай висит, и продолжает со следующего.
 */
import { readFileSync } from 'node:fs'
import type { CaseSet } from './cases.js'
import { evalCase } from './formulajs.js'

const [file, start] = process.argv.slice(2)
const { cases } = JSON.parse(readFileSync(file, 'utf8')) as CaseSet
for (let i = Number(start); i < cases.length; i++) {
  process.stdout.write(JSON.stringify({ i, r: evalCase(cases[i]) }) + '\n')
}
